import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

import os
import time
import xlwings as xw
from xlwings import Book
from typing import Dict, Literal
import xception as E, asyncio, pythoncom
from datetime import datetime

# SharePoint open-and-poll settings
_SP_POLL_INTERVAL = 1.0   # seconds between xw.books retries
_SP_POLL_TIMEOUT  = 30    # seconds total before giving up


class ActiveBook:

    def __init__(self):
        self.book_name: str = None

    async def set_target_book(self, book_name: str) -> Book:
        book_list = await get_working_book()
        if book_name in book_list:
            self.book_name = book_name
        else:
            raise E.NoBookFound(f'No Book named {book_name} found.')
        
    async def get_sheets(self) -> list:
        if self.book_name is None:
            raise E.NoBookSet('No Book Set')
        def sync_get_sheets():
            
            pythoncom.CoInitialize()
            try:
                book: Book = xw.books[self.book_name]
                return book.sheet_names
            except Exception as e:
                print(e)
                raise
            finally:
                pythoncom.CoUninitialize()
        return await asyncio.to_thread(sync_get_sheets)
        

        
    async def lookup(self, sheet_name: str, column: str, input_column: str, input_data: Dict[str, str], compare_mode: Literal['exact', 'str_value'], duplicate_key: Literal['allow', 'discard']) -> bool:
        def sync_lookup(col: str) -> bool:
            pythoncom.CoInitialize()
            try:
                book = xw.books[self.book_name]
                column = col.upper()
                sheet = book.sheets[sheet_name]
                endrow = get_last_row(sheet, column)

                for i in range(1, endrow + 1):
                    row = i
                    key_value = sheet[f'{column}{row}'].value
                    if compare_mode == 'str_value':
                        key_value = convert_excel_number(key_value)
                    if duplicate_key == 'allow':
                        i_data = input_data.get(key_value, '')
                    if duplicate_key == 'discard':
                        i_data = input_data.pop(key_value, '')
                    sheet[f'{input_column}{row}'].value = i_data
                return True
            except:
                raise
            finally:
                pythoncom.CoUninitialize()
        if self.book_name is None:
            raise E.NoBookSet('No Book Set')
        return await asyncio.to_thread(lambda col = column: sync_lookup(col))
        
    '''    async def read(self, sheet_name: str, column: str, st_row: int = 1, duplicate: Literal['allow', 'discard'] = 'discard', fetch_type: Literal['exact', 'string'] = 'string') -> tuple:
        async def add_to(holder_obj: list | set, add_data):
            if isinstance(holder_obj, list):
                holder_obj.append(add_data)
            if isinstance(holder_obj, set):
                holder_obj.add(add_data)

        
        
        if self.book_instance is None:
            raise E.NoBookSet('No Book Set')
        try:
            holder = [] if duplicate == 'allow' else set()
            column = column.upper()
            sheet = self.book_instance.sheets[sheet_name]
            lastrow = await get_last_row(sheet, column)
            st_row = st_row if isinstance(st_row, int) and st_row > 0 else 1
            for i in range(st_row, lastrow):
                values = sheet[f'{column}{i}'].value 
                if fetch_type == 'string':
                    values = await convert_excel_number(values)
                await add_to(holder, values)
            return holder
        except:
            raise'''

    async def read(self, sheet_name: str, column: str, st_row: int = 1, duplicate: Literal['allow', 'discard'] = 'discard', fetch_type: Literal['exact', 'string'] = 'string') -> tuple:
        
        book_name = self.book_name  # Store the name, not the COM object
        
        def _sync_read():
            pythoncom.CoInitialize()  # Initialize COM in this thread
            try:
                holder = [] if duplicate == 'allow' else set()
                column_upper = column.upper()
                book = xw.books[book_name]  # Re-acquire book in this thread
                sheet = book.sheets[sheet_name]
                lastrow = get_last_row(sheet, column)
                st = st_row if isinstance(st_row, int) and st_row > 0 else 1
                for i in range(st, lastrow):
                    values = sheet[f'{column_upper}{i}'].value 
                    if fetch_type == 'string':
                        if isinstance(values, float):
                            values = str(int(values))
                        else:
                            values = str(values)
                    if isinstance(holder, list):
                        holder.append(values)
                    else:
                        holder.add(values)
                return holder
            finally:
                pythoncom.CoUninitialize()  # Clean up COM
        
        if book_name is None:
            raise E.NoBookSet('No Book Set')

        return await asyncio.to_thread(_sync_read)

    async def simple_write(
        self,
        sheet_name: str,
        data: list[dict],
        fields: list[str],
        start_col: str,
        start_row: int,
        direction: Literal['vertical', 'horizontal'],
        write_headers: bool = False,
    ) -> int:
        """Write data rows to a sheet starting at start_col/start_row.

        Vertical: each data row → next Excel row; each field → column offset.
        Horizontal: each data row → next Excel column; each field → row offset.
        write_headers=True writes field names one step before the data start.
        Returns number of data rows written.
        """
        if self.book_name is None:
            raise E.NoBookSet('No Book Set')
        
        if write_headers:
            if direction == 'vertical':
                start_row = 2 if start_row <= 1 else start_row
            elif direction == 'horizontal':
                start_col = 'B' if start_col.upper() == 'A' else start_col.upper()

        book_name = self.book_name

        def _sync():
            pythoncom.CoInitialize()
            try:
                book = xw.books[book_name]
                sheet = book.sheets[sheet_name]
                start_col_num = col_letter_to_num(start_col)
                count = 0
                
                if direction == 'vertical':
                    if write_headers:
                        for i, field in enumerate(fields):
                            col_ltr = num_to_col_letter(start_col_num + i)
                            sheet[f'{col_ltr}{start_row - 1}'].value = field
                    if data:
                        data_2d = [[row_data.get(f) for f in fields] for row_data in data]
                        end_col = num_to_col_letter(start_col_num + len(fields) - 1)
                        end_row = start_row + len(data) - 1
                        sheet[f'{start_col}{start_row}:{end_col}{end_row}'].value = data_2d
                        count = len(data)

                elif direction == 'horizontal':
                    if write_headers:
                        prev_col = num_to_col_letter(start_col_num - 1)
                        for i, field in enumerate(fields):
                            sheet[f'{prev_col}{start_row + i}'].value = field
                    if data:
                        data_2d = [[row_data.get(f) for row_data in data] for f in fields]
                        end_col = num_to_col_letter(start_col_num + len(data) - 1)
                        end_row = start_row + len(fields) - 1
                        sheet[f'{start_col}{start_row}:{end_col}{end_row}'].value = data_2d
                        count = len(data)

                return count
            finally:
                pythoncom.CoUninitialize()

        return await asyncio.to_thread(_sync)

    async def lookup_write(
        self,
        sheet_name: str,
        data: list[dict],
        fields: list[str],
        key_field: str,
        lookup_ref: str,
        start_col: str,
        start_row: int,
        direction: Literal['vertical', 'horizontal'],
        write_headers: bool = False,
    ) -> int:
        """Match data rows against existing Excel content and write matched fields.

        Vertical: scan lookup_ref column; on match write fields at
                  (start_col+offset, matched_row).
        Horizontal: scan lookup_ref row; on match write fields at
                    (matched_col, start_row+offset).
        Excel float keys are normalised via convert_excel_number.
        write_headers=True writes field names one step before the data start.
        Returns number of matched rows/columns written.
        """
        if write_headers:
            if direction == 'vertical':
                start_row = 2 if start_row <= 1 else start_row
            elif direction == 'horizontal':
                start_col = 'B' if start_col.upper() == 'A' else start_col.upper()

        if self.book_name is None:
            raise E.NoBookSet('No Book Set')
        book_name = self.book_name

        def _sync():
            pythoncom.CoInitialize()
            try:
                book = xw.books[book_name]
                sheet = book.sheets[sheet_name]
                start_col_num = col_letter_to_num(start_col)
                count = 0
                write_fields = [f for f in fields if f != key_field]

                # Build normalised lookup map: key_string → data_row
                data_map: dict[str, dict] = {}
                for row_data in data:
                    raw = row_data.get(key_field, '')
                    key = convert_excel_number(raw) if isinstance(raw, (int, float)) else str(raw)
                    data_map[key] = row_data

                if direction == 'vertical':
                    if write_headers:
                        for i, field in enumerate(write_fields):
                            col_ltr = num_to_col_letter(start_col_num + i)
                            sheet[f'{col_ltr}{start_row - 1}'].value = field
                    end_row = get_last_row(sheet, lookup_ref)
                    col_values = sheet[f'{lookup_ref}1:{lookup_ref}{end_row}'].value
                    if not isinstance(col_values, list):
                        col_values = [col_values]
                    for r, cell_val in enumerate(col_values, start=1):
                        if cell_val is None:
                            continue
                        norm = convert_excel_number(cell_val) if isinstance(cell_val, (int, float)) else str(cell_val)
                        if norm in data_map:
                            row_data = data_map[norm]
                            for i, field in enumerate(write_fields):
                                col_ltr = num_to_col_letter(start_col_num + i)
                                sheet[f'{col_ltr}{r}'].value = row_data.get(field)
                            count += 1

                elif direction == 'horizontal':
                    if write_headers:
                        prev_col = num_to_col_letter(start_col_num - 1)
                        for i, field in enumerate(write_fields):
                            sheet[f'{prev_col}{start_row + i}'].value = field
                    end_col = get_last_col(sheet, int(lookup_ref))
                    last_col_ltr = num_to_col_letter(end_col)
                    row_values = sheet[f'A{lookup_ref}:{last_col_ltr}{lookup_ref}'].value
                    if not isinstance(row_values, list):
                        row_values = [row_values]
                    for c, cell_val in enumerate(row_values, start=1):
                        if cell_val is None:
                            continue
                        norm = convert_excel_number(cell_val) if isinstance(cell_val, (int, float)) else str(cell_val)
                        if norm in data_map:
                            col_ltr = num_to_col_letter(c)
                            row_data = data_map[norm]
                            for i, field in enumerate(write_fields):
                                sheet[f'{col_ltr}{start_row + i}'].value = row_data.get(field)
                            count += 1

                return count
            finally:
                pythoncom.CoUninitialize()

        return await asyncio.to_thread(_sync)


def convert_excel_number(number: float | int) -> str:
    if isinstance(number, float):
        return str(int(number))
    else:
        return str(number)

async def get_all_books_act() -> tuple:
    def _sync_get():
        pythoncom.CoInitialize()
        try:
            allbooks = tuple(bk.name for bk in xw.books)
            activate = xw.books.active.name
            return allbooks, activate
        except:
            raise
        finally:
            pythoncom.CoUninitialize()
    return await asyncio.to_thread(_sync_get)


async def get_working_book() -> tuple:
    def _sync_get():
        pythoncom.CoInitialize()
        try:
            return tuple(bk.name for bk in xw.books)
        finally:
            pythoncom.CoUninitialize()
    
    return await asyncio.to_thread(_sync_get)

def get_last_row(sheet: xw.Sheet, column: str) -> int:
    return sheet.api.Cells(sheet.api.Rows.Count, column).End(xw.constants.Direction.xlUp).Row

def col_letter_to_num(col: str) -> int:
    """'A'→1, 'B'→2, 'AA'→27 (1-based, matches xlwings column numbering)."""
    col = col.upper()
    result = 0
    for char in col:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result

def num_to_col_letter(n: int) -> str:
    """1→'A', 2→'B', 27→'AA' (inverse of col_letter_to_num)."""
    result = ''
    if n < 1:
        return 'A'
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(ord('A') + remainder) + result
    return result

def get_last_col(sheet: xw.Sheet, row: int) -> int:
    """Last used column number in a given row (1-based)."""
    return sheet.api.Cells(row, sheet.api.Columns.Count).End(xw.constants.Direction.xlToLeft).Column

async def get_active_book() -> str:
    def _sync_():
        pythoncom.CoInitialize()
        try:
            return xw.books.active.name
        except:
            raise
        finally:
            pythoncom.CoUninitialize()
    return await asyncio.to_thread(_sync_)

#async def test():
    #wb = await get_working_book()
    #print(wb)
    #ab = ActiveBook()

    #await ab.set_target_book('Book3')
    print(xw.books.active.name)
    #test_data = {'123' : 't', '126' : 'e', '128' : 's'}
    #await ab.lookup('Sheet1', 'C', 'E', test_data, 'str_value', 'discard')

    #a = await ab.read('Sheet1', 'C')
    #print(a)

#asyncio.run(test())


_CLOSE_POLL_INTERVAL = 2.0   # seconds between "is Excel still open?" checks
_CLOSE_POLL_TIMEOUT  = 7200  # give up after 2 h (user probably forgot to close)


async def wait_for_book_close(book_name: str) -> None:
    """Block until *book_name* disappears from xw.books (i.e. user closed it)."""
    def _sync():
        pythoncom.CoInitialize()
        try:
            deadline = time.monotonic() + _CLOSE_POLL_TIMEOUT
            while time.monotonic() < deadline:
                time.sleep(_CLOSE_POLL_INTERVAL)
                try:
                    xw.books[book_name]
                except Exception:
                    return  # book is gone
        finally:
            pythoncom.CoUninitialize()

    await asyncio.to_thread(_sync)


async def append_reschedule_row(
    request_str: str,
    username: str,
    lotid: str,
    db_figure: str,
    file_path: str,
    sheet_name: str,
) -> None:
    """
    Append one reschedule request row to the SharePoint Excel file.

    Row format (columns A-I):
        today | username | '' | lotid | '' | db_figure | '' | '' | request_str

    Opens the workbook by path if not already open in Excel.
    Uses get_last_row(sheet, 'A') to find the append position.
    """
    if request_str is None:
        raise ValueError('内容は必須です。')
    today = datetime.now().strftime('%Y/%m/%d')
    row_values = [today, username, '', lotid, '', db_figure, '', '', request_str]
    from urllib.parse import unquote
    book_name = unquote(Path(file_path).name.removesuffix('?web=1'))
    

    def _sync():
        pythoncom.CoInitialize()
        try:
            try:
                book = xw.books[book_name]
            except Exception:
                if not (file_path.startswith('\\\\') or (len(file_path) > 1 and file_path[1] == ':')):
                    # SharePoint URL: launch via Windows shell so Excel opens it,
                    # then poll until the workbook appears in xw.books.
                    os.startfile(file_path)
                    deadline = time.monotonic() + _SP_POLL_TIMEOUT
                    book = None
                    while time.monotonic() < deadline:
                        time.sleep(_SP_POLL_INTERVAL)
                        try:
                            book = xw.books[book_name]
                            break
                        except Exception:
                            continue
                    if book is None:
                        raise TimeoutError(
                            f'リスケ依頼表 {book_name} を {_SP_POLL_TIMEOUT}s 内に開くことができませんでした。リスケ依頼をコピーして手動で貼り付けてください。'
                        )
                else:
                    book = xw.Book(file_path)

            sheet = book.sheets[sheet_name]
            last = get_last_row(sheet, 'D') + 1
            for i, val in enumerate(row_values):
                sheet.cells(last, i + 4).value = val
        finally:
            pythoncom.CoUninitialize()

    await asyncio.to_thread(_sync)
