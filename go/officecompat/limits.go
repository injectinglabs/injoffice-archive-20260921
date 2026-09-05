package officecompat

import "fmt"

const (
	// These shared ceilings are public because the native mutation envelope
	// applies the exact package-byte limit before invoking a format handler.
	MaxPackageBytes         = 512 * 1024 * 1024
	MaxPackageParts         = 20_000
	MaxPartBytes            = 64 * 1024 * 1024
	MaxExpandedPackageBytes = 512 * 1024 * 1024
	MaxCompressionRatio     = 200

	defaultCompressionRatioSlack = 1 * 1024 * 1024
	defaultMaxXMLBytes           = 64 * 1024 * 1024
	defaultMaxXMLDepth           = 128
	defaultMaxXMLTokens          = 4_000_000
	defaultMaxXMLAttributes      = 4_000_000

	// Keep the established package inspector constant available to the
	// mutation hardening checks while explicit qualification envelopes use the
	// corresponding Limits field.
	compressionRatioSlack = defaultCompressionRatioSlack
)

// Limits bounds all package and XML work performed by the qualification
// helpers. A Limits value is fail-closed: every field must be positive except
// CompressionRatioSlack, which may be zero.
//
// DefaultLimits uses the largest current native Office extractor envelope so
// the format-neutral qualifier does not reject a package accepted by one of
// those engines merely because it is PPTX rather than DOCX or XLSX. Callers
// running smaller corpus shards should pass tighter limits explicitly.
type Limits struct {
	MaxPackageBytes       uint64
	MaxParts              int
	MaxPartBytes          uint64
	MaxExpandedBytes      uint64
	MaxCompressionRatio   uint64
	CompressionRatioSlack uint64
	MaxXMLBytes           uint64
	MaxXMLDepth           int
	MaxXMLTokens          uint64
	MaxXMLAttributes      int
}

// DefaultLimits returns a new copy of the qualification resource envelope.
func DefaultLimits() Limits {
	return Limits{
		MaxPackageBytes:       MaxPackageBytes,
		MaxParts:              MaxPackageParts,
		MaxPartBytes:          MaxPartBytes,
		MaxExpandedBytes:      MaxExpandedPackageBytes,
		MaxCompressionRatio:   MaxCompressionRatio,
		CompressionRatioSlack: defaultCompressionRatioSlack,
		MaxXMLBytes:           defaultMaxXMLBytes,
		MaxXMLDepth:           defaultMaxXMLDepth,
		MaxXMLTokens:          defaultMaxXMLTokens,
		MaxXMLAttributes:      defaultMaxXMLAttributes,
	}
}

func (limits Limits) validate() error {
	if limits.MaxPackageBytes == 0 || limits.MaxParts <= 0 || limits.MaxPartBytes == 0 || limits.MaxExpandedBytes == 0 || limits.MaxCompressionRatio == 0 || limits.MaxXMLBytes == 0 || limits.MaxXMLDepth <= 0 || limits.MaxXMLTokens == 0 || limits.MaxXMLAttributes <= 0 {
		return fmt.Errorf("officecompat: every resource limit except compression-ratio slack must be positive")
	}
	if limits.MaxPartBytes > limits.MaxExpandedBytes {
		return fmt.Errorf("officecompat: per-part limit exceeds expanded-package limit")
	}
	if limits.MaxXMLBytes > limits.MaxPartBytes {
		return fmt.Errorf("officecompat: XML-part limit exceeds per-part limit")
	}
	if limits.MaxPartBytes >= uint64(1<<63-1) || limits.MaxXMLBytes >= uint64(1<<63-1) {
		return fmt.Errorf("officecompat: stream limits exceed signed 64-bit reader capacity")
	}
	return nil
}
