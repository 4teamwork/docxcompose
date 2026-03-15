FROM python:3.14-alpine AS base

RUN addgroup -S -g 8080 docxcompose \
    && adduser -S -D -G docxcompose -u 8080 docxcompose

ENV PYTHONUNBUFFERED=1
WORKDIR /app

RUN echo "/app/lib/python3.14/site-packages/" > /usr/local/lib/python3.14/site-packages/app.pth \
 && apk add --no-cache \
    libxml2 \
    libxslt


FROM base AS builder

RUN apk add --no-cache \
    gcc \
    musl-dev \
    libxml2-dev \
    libxslt-dev \
    pipx

RUN pipx install poetry \
 && pipx inject poetry poetry-plugin-export

COPY pyproject.toml poetry.lock ./

RUN /root/.local/bin/poetry export -f requirements.txt --extras server --output requirements.txt \
 && pip install --no-cache-dir --no-warn-script-location --prefix ./ -r requirements.txt --no-binary lxml

COPY docxcompose docxcompose
COPY README.rst .
RUN pip install --no-cache-dir --prefix ./ .
RUN rm -rf docxcompose pyproject.toml poetry.lock poetry.toml README.rst requirements.txt


FROM base AS prod

COPY --from=builder /app /app
USER docxcompose
EXPOSE 8080
CMD ["/app/bin/docxcompose-server"]
