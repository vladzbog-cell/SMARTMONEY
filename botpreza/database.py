from typing import Any

import aiosqlite

DB_PATH = "orders.db"


async def init_db():
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("""
            CREATE TABLE IF NOT EXISTS orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                material TEXT,
                content_type TEXT,
                urgency TEXT,
                total_price INTEGER,
                file_format TEXT DEFAULT 'PDF',
                thread_id INTEGER,
                revisions_left INTEGER DEFAULT 2,
                status TEXT DEFAULT 'new',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        for col, typedef in [
            ("file_format", "TEXT DEFAULT 'PDF'"),
            ("thread_id", "INTEGER"),
        ]:
            try:
                await db.execute(f"ALTER TABLE orders ADD COLUMN {col} {typedef}")
            except Exception:
                pass
        await db.commit()


async def create_order(user_id: int, material: str, content_type: str,
                       urgency: str, total_price: int,
                       file_format: str = "PDF", thread_id: int | None = None) -> int:
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(
            """INSERT INTO orders
               (user_id, material, content_type, urgency, total_price, file_format, thread_id)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (user_id, material, content_type, urgency, total_price, file_format, thread_id),
        )
        await db.commit()
        return int(cursor.lastrowid or 0)


async def update_thread_id(order_id: int, thread_id: int):
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "UPDATE orders SET thread_id = ? WHERE id = ?", (thread_id, order_id)
        )
        await db.commit()


async def get_order(order_id: int) -> dict[str, Any] | None:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT * FROM orders WHERE id = ?", (order_id,)
        ) as cursor:
            row = await cursor.fetchone()
    return dict(row) if row else None


async def get_order_by_thread(thread_id: int) -> dict[str, Any] | None:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT * FROM orders WHERE thread_id = ?", (thread_id,)
        ) as cursor:
            row = await cursor.fetchone()
    return dict(row) if row else None


async def get_order_info(order_id: int) -> dict[str, Any] | None:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT user_id, status, revisions_left AS free_revisions, file_format "
            "FROM orders WHERE id = ?", (order_id,)
        ) as cursor:
            row = await cursor.fetchone()
    return dict(row) if row else None


async def check_and_use_revision(order_id: int) -> bool:
    async with aiosqlite.connect(DB_PATH) as db:
        async with db.execute(
            "SELECT revisions_left FROM orders WHERE id = ?", (order_id,)
        ) as cursor:
            row = await cursor.fetchone()
        if not row or row[0] <= 0:
            return False
        await db.execute(
            "UPDATE orders SET revisions_left = revisions_left - 1 WHERE id = ?",
            (order_id,),
        )
        await db.commit()
        return True


async def update_order_status(order_id: int, status: str):
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "UPDATE orders SET status = ? WHERE id = ?", (status, order_id)
        )
        await db.commit()
