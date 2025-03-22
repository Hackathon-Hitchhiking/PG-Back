up:
	uv run uvicorn app:app --reload --port 8000

.PHONY: migrate-rev
migrate-rev:
	@read -p "Enter the name of the revision: " name; \
	uv run alembic revision --autogenerate -m $$name

.PHONY: migrate-up
migrate-up:
	@read -p "Enter the revision to upgrade to: " rev; \
	uv run alembic upgrade $$rev

.PHONY: local
local:
	docker compose -f docker-compose.local.yml up

.PHONY: test
test:
	uv run pytest
