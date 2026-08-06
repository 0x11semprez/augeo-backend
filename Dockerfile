# syntax=docker/dockerfile:1

# ---- Build stage ----
FROM golang:1.22-alpine AS builder

WORKDIR /src

COPY go.mod go.sum ./
RUN go mod download

COPY . .
RUN CGO_ENABLED=0 GOOS=linux go build -ldflags="-s -w" -o /out/devis-api .

# ---- Runtime stage ----
# LibreOffice (headless) is required at runtime: internal/devis/convert.go
# shells out to `soffice` to convert the generated xlsx into a PDF.
# libreoffice-calc on Alpine/musl crashes on export (uno::RuntimeException,
# see https://gitlab.alpinelinux.org/alpine/aports/-/issues/9488), so the
# runtime stage uses glibc-based debian-slim instead; only the Go build
# stage above stays on Alpine.
FROM debian:bookworm-slim

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice-calc fonts-dejavu-core ca-certificates curl \
    && rm -rf /var/lib/apt/lists/* \
    && groupadd -r app && useradd -r -g app -m app

WORKDIR /app

COPY --from=builder /out/devis-api ./devis-api
COPY template/ ./template/
COPY data/ ./data/

RUN mkdir -p ./tmp && chown -R app:app /app

USER app

EXPOSE 8080

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD curl -sf http://localhost:8080/healthz || exit 1

ENTRYPOINT ["./devis-api"]
CMD ["-addr", ":8080", "-template", "./template/devis_template.xlsx", "-operateurs", "./data/operateurs.xlsx", "-workdir", "./tmp"]
