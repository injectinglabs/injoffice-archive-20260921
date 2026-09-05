// Command xlsxnative is a localhost helper that wraps xlsxpatch native extract
// and apply. It is not the hosted office server.
package main

import (
	"flag"
	"fmt"
	"io"
	"net"
	"net/http"
	"os"
	"strings"
	"time"

	"github.com/injectinglabs/injoffice/go/xlsxpatch/xlsxhttp"
)

const defaultAddr = "127.0.0.1:18765"

var usageText = `xlsxnative wraps go/xlsxpatch native XLSX extract/apply for local use.

Usage:
  xlsxnative extract [input.xlsx]
  xlsxnative apply original.xlsx payload.json
  xlsxnative serve [--addr 127.0.0.1:18765]

extract reads an XLSX (path or stdin) and writes native v2 JSON to stdout.
apply reads original XLSX bytes plus a native mutation payload and writes the
saved XLSX to stdout. The payload is the bounded JSON object accepted by
ApplyNativeWorkbookMutationPayloadV1 (expected_revision is rev:<digest>).
The outer CAS sha256:<digest> is taken from the original bytes.

serve binds a loopback-only helper (no auth):
  POST /v1/xlsx/extract     raw XLSX body or multipart file field
  POST /v1/xlsx/mutations   multipart original + payload + expected_revision

The playground native preview expects this helper:
  go run ./go/xlsxpatch/cmd/xlsxnative serve
`

func main() {
	os.Exit(run(os.Args[1:], os.Stdin, os.Stdout, os.Stderr))
}

func run(args []string, stdin io.Reader, stdout, stderr io.Writer) int {
	if len(args) == 0 || args[0] == "-h" || args[0] == "--help" || args[0] == "help" {
		fmt.Fprint(stderr, usageText)
		if len(args) == 0 {
			return 2
		}
		return 0
	}
	switch args[0] {
	case "extract":
		data, err := readInput(args[1:], stdin)
		if err != nil {
			fmt.Fprintln(stderr, "xlsxnative extract:", err)
			return 1
		}
		encoded, err := xlsxhttp.ExtractNativeJSON(data)
		if err != nil {
			fmt.Fprintln(stderr, "xlsxnative extract:", err)
			return 1
		}
		if _, err := stdout.Write(append(encoded, '\n')); err != nil {
			fmt.Fprintln(stderr, "xlsxnative extract:", err)
			return 1
		}
		return 0
	case "apply":
		if len(args) != 3 {
			fmt.Fprintln(stderr, "xlsxnative apply: usage: xlsxnative apply original.xlsx payload.json")
			return 2
		}
		original, err := os.ReadFile(args[1])
		if err != nil {
			fmt.Fprintln(stderr, "xlsxnative apply:", err)
			return 1
		}
		payload, err := os.ReadFile(args[2])
		if err != nil {
			fmt.Fprintln(stderr, "xlsxnative apply:", err)
			return 1
		}
		produced, err := xlsxhttp.ApplyNativeMutation(original, payload, "")
		if err != nil {
			fmt.Fprintln(stderr, "xlsxnative apply:", err)
			return 1
		}
		if _, err := stdout.Write(produced); err != nil {
			fmt.Fprintln(stderr, "xlsxnative apply:", err)
			return 1
		}
		return 0
	case "serve":
		fs := flag.NewFlagSet("serve", flag.ContinueOnError)
		fs.SetOutput(stderr)
		addr := fs.String("addr", defaultAddr, "loopback host:port")
		if err := fs.Parse(args[1:]); err != nil {
			return 2
		}
		if err := requireLoopbackAddr(*addr); err != nil {
			fmt.Fprintln(stderr, err)
			return 1
		}
		fmt.Fprintf(stderr, "xlsxnative: listening on http://%s (localhost helper, no auth)\n", *addr)
		server := &http.Server{
			Addr:              *addr,
			Handler:           xlsxhttp.NewHandler(nil),
			ReadHeaderTimeout: 5 * time.Second,
			ReadTimeout:       60 * time.Second,
			WriteTimeout:      60 * time.Second,
			MaxHeaderBytes:    1 << 16,
		}
		if err := server.ListenAndServe(); err != nil {
			fmt.Fprintln(stderr, "xlsxnative serve:", err)
			return 1
		}
		return 0
	default:
		fmt.Fprintf(stderr, "xlsxnative: unknown command %q\n\n%s", args[0], usageText)
		return 2
	}
}

func requireLoopbackAddr(addr string) error {
	host, port, err := net.SplitHostPort(addr)
	if err != nil {
		return fmt.Errorf("xlsxnative: serve: --addr must be host:port: %w", err)
	}
	if port == "" {
		return fmt.Errorf("xlsxnative: serve: --addr must include a port")
	}
	if strings.EqualFold(host, "localhost") {
		return nil
	}
	ip := net.ParseIP(host)
	if ip == nil || !ip.IsLoopback() {
		return fmt.Errorf("xlsxnative: serve: --addr must bind to localhost only")
	}
	return nil
}

func readInput(args []string, stdin io.Reader) ([]byte, error) {
	if len(args) == 0 || args[0] == "-" {
		return io.ReadAll(stdin)
	}
	return os.ReadFile(args[0])
}
