# Build from the repository root:
#   docker build -f go/injoffice-server/Dockerfile .
FROM golang:1.23-alpine AS build
WORKDIR /src
COPY go/xlsxpatch ./go/xlsxpatch
COPY go/docxpatch ./go/docxpatch
COPY go/pptxpatch ./go/pptxpatch
COPY go/collab ./go/collab
COPY go/injoffice-server ./go/injoffice-server
WORKDIR /src/go/injoffice-server
RUN CGO_ENABLED=0 go build -trimpath -ldflags="-s -w" -o /out/injoffice-server ./cmd/injoffice-server

FROM alpine:3.21
RUN adduser -D -H -u 65532 nonroot \
	&& mkdir -p /var/lib/injoffice/artifacts \
	&& chown nonroot:nonroot /var/lib/injoffice/artifacts
USER nonroot
COPY --from=build /out/injoffice-server /usr/local/bin/injoffice-server
ENV INJOFFICE_ADDR=0.0.0.0:18765
ENV INJOFFICE_ARTIFACTS=/var/lib/injoffice/artifacts
EXPOSE 18765
ENTRYPOINT ["/usr/local/bin/injoffice-server"]
