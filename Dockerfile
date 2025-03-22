# Build stage
FROM python:3.12-slim-bookworm AS builder

# Install uv
COPY --from=ghcr.io/astral-sh/uv:latest /uv /bin/uv

# Set working directory
WORKDIR /app

# Copy dependency files
COPY pyproject.toml uv.lock ./

# Install dependencies
RUN --mount=type=cache,target=/root/.cache/uv \
    uv pip install --system --no-deps -r pyproject.toml

# Copy project files
COPY . .

# Install project
RUN uv pip install --system -e .

# Final stage
FROM python:3.12-slim-bookworm

# Copy installed packages and project files
COPY --from=builder /usr/local/lib/python3.12/site-packages /usr/local/lib/python3.12/site-packages
COPY --from=builder /app /app

# Set working directory
WORKDIR /app

# Run the application
CMD ["python", "app.py"]
