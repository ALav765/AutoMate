import os
import asyncpg
from dotenv import load_dotenv

load_dotenv()

_pool = None

async def get_pool():
    global _pool
    if _pool is None:
        _pool = await asyncpg.create_pool(
            dsn=os.environ.get("DATABASE_URL", "postgresql://localhost:5432/supply_plan"),
            min_size=1,
            max_size=5,
        )
    return _pool

async def close_pool():
    global _pool
    if _pool is not None:
        await _pool.close()
        _pool = None
