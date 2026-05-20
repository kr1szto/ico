# IČO Lookup

FastAPI web wrapper for supplier lookup by Slovak IČO.

## Deployment status

The repository now contains executable deployment files:

- `app.py`
- `requirements.txt`
- `Dockerfile`
- `docker-compose.yaml`

Railway can build the Docker image and the container listens on `${PORT:-8000}`.

The scraper implementation is still required separately. Add `main.py` with:

```python
def scrape_subject(ico: str) -> dict:
    ...
```

Until `main.py` exists, the app can start but `/lookup` will return a configuration error.
