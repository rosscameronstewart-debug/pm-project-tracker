import sqlite3
import tempfile
import threading
import unittest
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from unittest.mock import patch
import app


class DatabaseSettingsTests(unittest.TestCase):
    def test_connection_timeout_and_durability(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(app,'DB_PATH',Path(directory)/'settings.sqlite3'):
            con=app.db()
            try:
                self.assertEqual(con.execute('PRAGMA busy_timeout').fetchone()[0],15000)
                self.assertEqual(con.execute('PRAGMA synchronous').fetchone()[0],2)
                self.assertEqual(con.execute('PRAGMA foreign_keys').fetchone()[0],1)
            finally:con.close()

    @unittest.skipIf(sqlite3.sqlite_version_info < (3,51,3),'Requires approved patched runtime')
    def test_wal_reads_parallel_writes_numbers_and_backup(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(app,'DB_PATH',Path(directory)/'wal.sqlite3'):
            reader=app.db()
            self.assertEqual(reader.execute('PRAGMA journal_mode=WAL').fetchone()[0],'wal')
            reader.execute('CREATE TABLE purchase_orders (po_number TEXT UNIQUE)')
            reader.commit()
            reader.execute('BEGIN')
            self.assertEqual(reader.execute('SELECT count(*) FROM purchase_orders').fetchone()[0],0)
            def write_number(_):
                con=app.db()
                try:
                    with con:
                        number=app.next_po_number(con,'TEST')
                        con.execute('INSERT INTO purchase_orders VALUES (?)',(number,))
                    return number
                finally:con.close()
            with ThreadPoolExecutor(max_workers=8) as pool:
                numbers=list(pool.map(write_number,range(24)))
            self.assertEqual(len(set(numbers)),24)
            self.assertEqual(reader.execute('SELECT count(*) FROM purchase_orders').fetchone()[0],0)
            reader.rollback()
            self.assertEqual(reader.execute('SELECT count(*) FROM purchase_orders').fetchone()[0],24)
            # A blocked writer waits, then succeeds after the first releases its lock.
            writer=app.db();writer.execute('BEGIN IMMEDIATE')
            started=threading.Event()
            def queued_write():
                started.set()
                return write_number(0)
            with ThreadPoolExecutor(max_workers=1) as pool:
                pending=pool.submit(queued_write)
                self.assertTrue(started.wait(2))
                writer.commit();writer.close()
                self.assertTrue(pending.result(timeout=10))
            target=sqlite3.connect(Path(directory)/'backup.sqlite3')
            reader.backup(target)
            self.assertEqual(target.execute('PRAGMA integrity_check').fetchone()[0],'ok')
            self.assertEqual(target.execute('SELECT count(*) FROM purchase_orders').fetchone()[0],25)
            target.close();reader.close()


if __name__=='__main__':unittest.main()
