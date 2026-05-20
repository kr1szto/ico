# IČO Lookup

FastAPI web wrapper for supplier lookup by Slovak IČO.

## Deployment status

The repository now contains executable deployment files:

- `app.py`
- `main.py`
- `requirements.txt`
- `Dockerfile`
- `docker-compose.yaml`
- `fly.toml`

Railway or Fly.io can build the Docker image, install Chromium/ChromeDriver, and run the FastAPI app on `${PORT:-8000}`.

## Runtime modes

- Railway, Fly.io, or single-container Docker: uses local Chromium through ChromeDriver.
- Docker Compose: uses the `selenium` sidecar because `SELENIUM_URL` is set to `http://selenium:4444/wd/hub`.

## Fly.io

Deploy from the repository root with:

```bash
fly deploy
```

The included `fly.toml` uses app name `kr1szto-ico`, exposes the Dockerfile's `8000` port, and starts with a 1 GB shared CPU machine. Increase memory to 2 GB if Chromium exits under load.

External registry pages can still change markup or block automation; scraper failures should be treated as runtime risks, not deployment wiring issues.
