from oracledb import connect_async #,create_pool_async, AsyncConnectionPool, 
from sql_dict import sql_dict as SD, predefined
from datetime import datetime, timedelta
import re, sys, os, socket, shelve, xception as E, pandas as pd
from collections import OrderedDict, Counter
from typing import List, Tuple, Self
#from concurrent.futures import ThreadPoolExecutor
from aggridsetting import generate_col_settings
from typing import Dict, List, Tuple, Self
import pyperclip
from html.parser import HTMLParser
from nicegui import ui


#C:\app\oracle\product\19.3.0\client_1\network\admin#

def connection_wrapper(func):
    async def check_connection(*args, **kwargs):
        host = socket.gethostname()
        ip = socket.gethostbyname(host)
        if not ip.startswith("172") or ip.startswith("127.0.0"):
            raise E.NoConnectionExc()
        output = await func(*args, **kwargs)
        return output
    return check_connection


_CONN_MAX_AGE    = timedelta(minutes=25)   # proactive reconnect threshold
_CALL_TIMEOUT_MS = 120_000                 # max ms for any single DB operation


class Odbconn_Cursor:
    __slots__ = ["lastsearch", "pool_create", "executor", "_connect_time"]

    MODE_BOK = {
    "b":("prc",("LotID","UserCode","ProductCode","CurrentProcess","Status","OnHoldReason","LastUpdatedDate")),
    "r":("rls",("ParentID","ParentSAP","ChildID","ChildSAP")),
    "t":("rls-t",("NodeID",)),
    "i":("ipr",("LotID","UserCode","ProductCode","CurrentProcess","LotState","OnHoldReason","LastUpdatedDate","InProgress","Restriction","flowName")),
    "v":("vie",("LotID","ProcessName","Start","Finish","Resource","Status")),
    "w":("wip",("WIP-I",)),
    "e":("wip-c",("WIP-I",)),
    "s":("srn",("LotID","SerialNumber")),
    "h":("rcd",("RecordsCluster",)),
    "d":("dbr",("DB",)),
    "c":("dbr-c",("DB",)),
    "y":("ybg-l",("YellowBagCluster",)),
    "l":("ybg-d",("YellowBagCluster",)),
    "a":("dbs", ("DB",)),
    "p":("cpn", ("DB", "SAPCode", "CompoName", "ComponentCode", "Count", "StrName")),
    "f":("flo", ("flowitem",)),
    "q":("she-e", ("existing",)),
    "U":("she-u", ("updating",)),
    "I":("she-i", ("history",)),
    "A":("she-a", ("all_data",)),
    "D":("data-1", ("dataviewer",)),
    "K":("data-2", ("dataviewer",)),
    "F":("data-3", ("dataviewer",)),
    "S":("data-5", ("dataviewer",)),
    "P":("pov", ("pov",))
    }
    # single character flag : (from and where flag,(select flag))
    #OR_CONFIGURE = r"\\sflise02\伊勢原２\接合・セラミック部\業務課\2.業務G\Apps\ODBtns\tnsnames.ora"

    def __init__(self):
        self.lastsearch = None #tuple(sql stirng, last searched time)
        #self.executor = ThreadPoolExecutor(max_workers= 2)
        self.pool_create = None
        self._connect_time = None
        #defaults.config_dir = Odbconn_Cursor.OR_CONFIGURE

    @connection_wrapper
    async def p_connect(self):
        # Proactive reconnect if connection exceeds max age
        if self.pool_create is not None and self._connect_time is not None:
            if datetime.now() - self._connect_time > _CONN_MAX_AGE:
                try:
                    await self.pool_create.close()
                except Exception:
                    pass
                self.pool_create = None
                self._connect_time = None

        if self.pool_create is None:
            try:
                self.pool_create = await connect_async()
                self.pool_create.call_timeout = _CALL_TIMEOUT_MS
                self._connect_time = datetime.now()
            except Exception as e:
                raise E.NoDBConnectionExc(f'Error occurred while creating database connection pool: {str(e)}')

    async def ensure_connection(self) -> bool:
        """Warm up DB connection proactively. Safe to fire-and-forget."""
        print('Ensuring database connection...')
        try:
            await self.p_connect()
            print('Database connection is healthy.')
            return True
        except (E.NoConnectionExc, E.NoDBConnectionExc):
            print('Database connection is not available.')
            return False
        except Exception as e:
            print(f'Unexpected error occurred: {e}')
            return False

    def q_execute(self, sql:str, parameters: list | dict | tuple | None = None, outaspd: bool = False) -> pd.DataFrame | List[Tuple[str]]:
        #print('Trying to acquire connection')
        with self.pool_create as connection:
            #print('Connection acquired.')
            #try:
            #    retrive = await connection.fetchall(sql, parameters=parameters)
            #    print(f'retrive corotione = {retrive}')
            if outaspd:
                try:
                    retrive = pd.read_sql(sql, connection)
                except Exception as e:
                    raise E.SQL_ErrorExc(e)
                return retrive
                
            with connection.cursor() as cursor:
                #print(f'module: 68 -> {parameters}')
                #print(sql)
                try:
                    retrive = cursor.execute(sql, parameters=parameters).fetchall()
                except Exception as e:
                    raise E.SQL_ErrorExc(e)
                
        return retrive
    
    async def asy_q_execute(self, sql:str, parameters: list | dict | tuple | None = None, outaspd: bool = False, commit: bool = False):
        #import time
        #start_of_query_execution = time.perf_counter()
        await self.p_connect()
        #print(f'Connection acquired in {time.perf_counter() - start_of_query_execution:.2f} seconds')
        connection = self.pool_create
        #start_of_query_execution = time.perf_counter()
        with connection.cursor() as c:
            c.arraysize = 1000
            await c.execute(sql, parameters=parameters)
            #print(sql)
            #print(f'cQuery executed in {time.perf_counter() - start_of_query_execution:.2f} seconds')
            if commit:
                await connection.commit()
                return None
            else:
                try:
                    #start_of_query_execution = time.perf_counter()
                    retrive = await c.fetchall()
                    #print(f'cQuery fetchall executed in {time.perf_counter() - start_of_query_execution:.2f} seconds')
                except Exception as e:
                    raise E.SQL_ErrorExc(e)

                if outaspd:
                    try:
                        #start_of_query_execution = time.perf_counter()
                        headrow = [head[0] for head in c.description]
                        retrive = pd.DataFrame(retrive, columns=headrow)
                        #print(f'cQuery DataFrame conversion executed in {time.perf_counter() - start_of_query_execution:.2f} seconds')
                    except Exception as e:
                        raise E.InternalErrorExc(e)

        return retrive


    
    async def mod_execute(self, batchsql: str, values: List[str], recurlevel:int = 0) -> bool:
        await self.p_connect()
        truncated_first = values
        
        if len(values) > 999:
            truncated_first = values[:999]
            truncated_second = values[999:]
            newrecurlevel = recurlevel + 1
            recur_end_as = await self.mod_execute(batchsql, truncated_second, newrecurlevel)
            if not recur_end_as:
                raise Exception('Error happened during SQL operation.')
        connection = self.pool_create
        with connection.cursor() as cursor:
            cursor.setinputsizes()
            try:
                await cursor.executemany(batchsql, truncated_first)
            except Exception as e:
                print(e)
                return False
        if recurlevel == 0:
            try:
                await connection.commit()

                return True
            except Exception as e:
                return False
        else:

            return True
        
    async def sql_plain(self, sql, try_count: int = 0, outaspd: bool = False, commit: bool = False):

        search_time = await self.frequency_limiter(sql)
        result = None
        retry_limit = 3
        while try_count < retry_limit:
            #print(f'********************\nTry Count = {try_count}')
            try:

                result = await self.asy_q_execute(sql, outaspd=outaspd, commit= commit)
                #print(f'Got result = \n\n{result}')

                if outaspd:
                    if result.empty:
                        raise E.NoDataExc()
                else:
                    if not result:
                        raise E.NoDataExc()

                return result, search_time

                    
            except E.NoDataExc:
                return [], search_time
            except Exception as error:
                error_str = str(error)
                if 'DPY-4011' in error_str or 'DPY-4024' in error_str or 'Errno 11001' in error_str:
                    self.lastsearch = None
                    print(f'#############################\nConnection Closed: \n {error_str}\nConnection Attempting Reconnect...\nAttempt {try_count + 1}')
                    try:
                        if self.pool_create is not None:
                            await self.pool_create.close()
                            self.pool_create = None
                            self._connect_time = None
                        await self.p_connect()
                    except E.NoDBConnectionExc as e:
                        print(f'NoDBConnectionExc: {str(e)}')
                    except Exception as eelse:
                        print(f'Unexpected error: {str(eelse)}')
                    finally:
                        try_count += 1
                        print('\n#############End of Connection Attempt Log##################')
                    continue
                else:
                    raise E.SQL_ErrorExc(f'###########################SQL Error#######################\n{error_str}\n############################################\n')

        raise E.NoDBConnectionExc(str(error_str))
    
    async def sql_plain_modify(self, sql_list:Dict[str, List[str]]) -> bool:
        
        result = None
        for sql, valuelist in sql_list.items():
            #print(f'valuelist = \n{valuelist}')
            try:
                result = await self.mod_execute(sql, valuelist)
                if result:
                    return result
                else:
                    raise Exception('Error Happened in SQL operations.')
            except Exception as e:
                raise


    async def search(self, lid, mode, **kwargs):
        '''
        Part of code needs to be altered when new search pattern is added:
        1.module.Obdconn_Cursor.search:
         add new pattern to mode_bok: {mode flag:(modeflag, tuple(column and attributes name))}
        2.gui.TreeView.column_dict:
         add new column pattern {mode name:{column name:column width}}
        '''

        #is the 3 letter flag nesessary? its not used anywhere, check later: its required as a parameter to passed to LotObject
        #for index, code in enumerate(lid):
        #    if "'" not in lid[index]:
        #        lid[index] = f"'{lid[index]}'"
        #change the behaviour of the way id string is assemabled, the original way will affect the passed in parameter outside
        #the function

        id_string = tuple(f"'{id}'" for id in lid)
        if len(id_string) > 1:
            lotid_str = ",".join(id_string)
        else:
            lotid_str = id_string[0] if id_string else ''
        try:
            obj = LotObject(lid=lotid_str, cmd=Odbconn_Cursor.MODE_BOK[mode][0], att=Odbconn_Cursor.MODE_BOK[mode][1], cond=kwargs["cond"], plant=kwargs["plant"])
            string_ = SD(obj)
            #print(f'Generated SQL string = {string_}')
        #print(string_)
            import time
            #start_of_query_execution = time.perf_counter()
            cursor_contents, search_time = await self.sql_plain(string_, outaspd=kwargs.get('outaspd', False))
            #print(f'Query SQL Plain execution completed in {time.perf_counter() - start_of_query_execution:.2f} seconds')
        except:
            raise
        return cursor_contents, search_time

    async def frequency_limiter(self, sql):
        search_time ,search_time_st = time_now()
        if self.lastsearch is None:
            self.lastsearch = (sql, search_time)
            return search_time_st
        delta = (search_time - self.lastsearch[1]) < timedelta(minutes=3)
        if sql == self.lastsearch[0] and delta:
            raise E.FrequentExc()
        self.lastsearch = (sql, search_time)
        return search_time_st

  
    


class LotObject:
    #__slots__ = ["lotid", "cmd", "attributes", "operator", "cond", "plant"]
    RULE_BOOK = {
    "prc":("ProcessCheck","ProcessRelated","ProcessRelated"),
    "rls":("RelatedID","RelatedSearch","RelatedSearch"),
    "rls-t":("RelatedID_Tree","RelatedSearch_tree","RelatedSearch_tree"),
    "ipr":("ProcessCheck","ProcessRelated","ProcessRelated"),
    "vie":("Schedule", "ScheduleView", "ScheduleView"),
    "wip":("WIP-I", "WIP-I", "WIP-I"),
    "wip-c":("WIP-I", "WIP-I", "WIP-IC"),
    "rcd":("Records","Records","Records"),
    "dbr":("RelatedID", "DBRelation", "DBRelation"),
    "dbr-c":("RelatedID", "DBRelation", "DBRelation_C"),
    "srn":("ProcessCheck","Serial","ProcessRelated"),
    "ybg-l":("YellowBag", "YellowBag", "YellowBag_Lot"),
    "ybg-d":("YellowBag", "YellowBag", "YellowBag_DB"),
    "dbs" : ("DBSearch", "DBSearch", "DBSearch"),
    "cpn" : ("Components", "Components", "Components"),
    "flo" : ("Flow", "Flow", "Flow"),
    "she-e" : ("ShelfExisting", "existing", "existing"),
    "she-u" : ("ShelfUpdating", 'updating', 'updating'),
    "she-i" : ("HistoryInsert", 'updating', 'history'),
    "she-a" : ("ShelfExisting", "all_data", "all_data"),
    "data-1" : ("DataViewerUni", "YellowBag", "dataview_none"),
    "data-2" : ("DataViewerUni", "ksystem", "dataview_none"),
    "data-3" : ("DataViewerUni", "product_inventory", "dataview_none"),
    "data-5" : ("DataViewerUni", "ScheduleView", "schedule_view_all"),
    "pov" : ("ProcessOverview", "pov", "pov")
    }
    __slots__ = ["lotid", "cmd", "attributes", "operator", "cond", "plant", "datas"]
    def __init__(self, **kwargs):
        self.lotid = kwargs["lid"]
        self.cmd = kwargs["cmd"]
        self.attributes = kwargs["att"]
        self.operator = LotObject.RULE_BOOK[self.cmd]
        self.cond = kwargs["cond"]
        self.plant = kwargs.get("plant", "5501")
        self.datas: dict = kwargs.get("datas", {})

        


OC = Odbconn_Cursor()
#asyncio.run(OC.p_connect())
LastSearchedCache = OrderedDict()

async def update_lead_time(db_code: str) -> bool:
    sql_dict = {
        '''
        INSERT INTO PROMAN_AVG_LEADTIME (DB, PROCESS_CODE, LEADTIME, CREATED_DATE) 
        SELECT 
            NHK図番, 
            工程コード, 
            ROUND(Avg(リードタイム), 2), 
            CURRENT_DATE 
            FROM 
                LV23_TEST.V_工程別指図リードタイム_完成
            WHERE (NHK図番= :DB) 
            GROUP BY 
                NHK図番, 
                工程コード
        ''' :
        [(db_code, )]
    }
    try:
        result = await OC.sql_plain_modify(sql_dict)
    except Exception as e:
        print(e)
        return False
    return True


async def run_sql_plain(commit: bool = False):
    import traceback
    sql = input('Input SQL >>')
    try:
        result = await OC.sql_plain(sql, commit= commit)
    except Exception as e:
        print(f'EXCEPTION:\n{e}')
        traceback.print_exc()

    print(f'############################\n##########RESULT############\n\nType = {type(result)}\n\nLength = {len(result)}\n\nContents = {result}')
    print('\n\nEND')

def display(object):
    '''
    This function is for data table display in command line enviroment, this is not necessary for gui.
    When it comes to command line interface development one day, this will be used to fetch the result as string
    '''
    from os import system
    from datarender import Table
    os.system("cls")
    header = {
    "b":["LotID","UserCode","ProductCode","CurrentProcess","WIPStatus","ProdStatus","LastUpdatedDate"]
    }
    output = Table(object,header["b"]).table()
    return output

def dbp(*args):
    for i, arg in enumerate(args):
        print(f"Arg{i+1} = {arg}")

def attrnamelower(name):
    try:
        name = name.lower()
    except:
        name = name
    return name

def delaydetect(now, before, **kwargs):
    timeformat = "%Y-%m-%d %H:%M"
    time_obj= now[7]
    delta = (datetime.now() - time_obj) > timedelta(minutes=60)
    status_coll = ("投入", "受入", "完成", "梱包", "支給", "部品出庫")
    if any(status in now[3] and status in before[3] for status in status_coll):
        return 0
    if "保留" in now[4]:
        return 1
    if now[3]==before[3] and delta and now[8]==before[8]:
        return 2
    else:
        return 0

def time_now():
    timestamp = datetime.now()
    return timestamp, timestamp.strftime('%Y-%m-%d %H:%M')

def notice_string(groups):
    string = "Update on products under monitoring:\n"
    for name in groups.keys():
        group = groups[name]
        uppername = name.upper()
        if len(group) == 1:
            string += f"\tLot id: {group[0]} is {uppername}.\n"
        else:
            string += f"\t{len(group)} lot id : {uppername}.\n"
    return string

def vari_split(text):
    for seperator in (",./"):
        if seperator in text:
            splited = tuple(item.strip() for item in text.split(seperator))
            return splited
    else:
        return text.split()


def regex_split(text: str, mode: str) -> List[str]:
    text = text.upper()
    pattern_dict = {"lot":r"42\d{8}|[KNS](?:\d|-){6,10}",
                    "prodcode":r'DB\d{4,5}(?:-(?:W\d{4}|B\d{4}|\d{4,5}))?',
                    "like":r"\bDB(?=[\d\-WB\*]{0,10}\*(?!\w))(?![\d\-WB\*]*\*[\d\-WB\*]*\*)(?![\d\-WB\*]*-[\d\-WB\*]*-)(?![\d\-WB\*]*[WB][\d\-WB\*]*[WB])[\d\-WB\*]{0,10}\*"}
    pattern = pattern_dict[mode]
    val_splited = list(set(re.findall(pattern, text)))
    #print(val_splited)
    if mode == "like":
        newlist = []
        for key in val_splited:
            newkey = key.replace('*', '%')
            newlist.append(newkey)
        val_splited = newlist
    if len(val_splited) < 1000:
        return val_splited
    return val_splited[:1000]


def copytoclipboard(copied: List[dict]) -> bool:
    concat_copied = '\n'.join(tuple('\t'.join(map(str, tuple(item.values()))) for item in copied))# + '\n'
    #print('copied!!')
    pyperclip.copy("")
    pyperclip.copy(concat_copied)
    if pyperclip.paste() == concat_copied:
        return True

def singlecopy(data: str | int) -> str:
    if not isinstance(data, (str, int)) or not data:
        return 0
    pyperclip.copy("")
    pyperclip.copy(data)
    if pyperclip.paste() == data and data:
        return 1

def htmltable(contents):
    table_head = '<table style="border-collapse: collapse; width: 100%;">\n'
    td_head = '<td style="border: 1px solid #dddddd; text-align: left; padding: 8px;">'
    if contents:
        table = '<!DOCTYPE html>\n<html>\n<body>\n'
        table += table_head
        for row in contents:
            record = "<tr>\n"
            record += "\n".join(tuple(f"{td_head}{r}</td>" for r in row.split("\t")))+"\n"
            record += "</tr>\n"
            table += record
        table += "</table>\n</body>\n</html>"
        return table


def checkfrozen(relative_path):
    if getattr(sys, "frozen", False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


async def wrapped_sql_search(keys: List[str] | Tuple[str], mode: str = "i", **kwargs):
    '''
    - keys: A string containing the search keys.(Used with SQL query)
    - mode: A string indicating the search mode, defaulting to "i".
        Avalaible key refer to Odbconn_Cursor class.
    - kwargs: Additional keyword arguments.
    '''
    #print(f'wrapped_sql_search = {kwargs.get('outaspd')}')

    cond = kwargs.get("cond", (0,0))
    if keys:
        joined_key = "".join(keys)
        '''if mode is None or mode not in "brvwshdc":
            mode = "i"'''
    else:
        joined_key = ""
    try:
        result, search_time = await OC.search(keys, mode, cond=cond, plant=kwargs["plant"], outaspd=kwargs.get('outaspd', False))
    except E.FrequentExc:

        return LastSearchedCache.get(joined_key, ("NoCacheAvaliable",))
    except:
        raise
    LastSearchedCache[joined_key] = result
    #print(f'Cache Updated: {joined_key} stored in cache at {search_time}. Cache size: {len(LastSearchedCache)}')
    clear_cache()
    #con_timeout(kwargs["afterid"])

    return result

    
async def fast_query(query_key:str, parameters: list | tuple | dict, outaspd: bool = False, commit: bool = False):
        try:
            sql = predefined(query_key)
            print(f'sql = {sql}')
        except Exception as e:
            print(f'Error retrieving SQL for query_key "{query_key}": {str(e)}')
            return None
        if query_key[:5] != 'NOKEY' and len(parameters) == 0:
            return None
        if sql is not None:
            try:
                result = await OC.asy_q_execute(sql, parameters, outaspd, commit=commit)
                #print(f'fast_query executed for query_key "{query_key}" with parameters {parameters}. Result: {result}')
                return result
            except Exception as e:
                print(f'Error in fast_query [{query_key}]: {str(e)}')
                return None
        else:
            return None



def add_on_click(cur_tree, to_tree, timestamp, key_cache, builder, plant_code):
    cleaned_keys = tuple(key for key in cur_tree.get_value(0) if key_cache is None or key not in key_cache)
    last = 0 if not key_cache else len(key_cache)
    if cleaned_keys:
        new_cursor = to_tree.wrapped_insert(wrapped_sql_search(timestamp, cleaned_keys, afterid=builder, plant=plant_code), tags=True)
        key_cache.update(new_cursor)
        to_tree.contextmenu_init()
        return
    raise E.NoKeyExc()

def lotid_range(lotid,ran=None,end=None):
    pattern = r"42\d{8}|[KNS](?:\d|-){6,10}"
    if not re.match(pattern, lotid):
        raise E.InvalidInputExc()
    if ran is not None:
        result = (int(lotid)+num for num in range(ran))
    if end is not None:
        result = (id for id in range(int(lotid), int(end)+1))
    return tuple(result)

def clear_cache():
    if len(LastSearchedCache) > 6:
        LastSearchedCache.popitem(last=False)

def close_pool():
    try:
        OC.pool_create.close()
        return
    except:
        OC.pool_create.close(force=True)
    finally:
        return


def shelve_op_ext(key):
    with shelve.open("config") as db:
        result = db[key]
    return result

def shelve_op_ext_m(key_arr):
    result_dict = {}
    with shelve.open("config") as db:
        for key in key_arr:
            result_dict[key] = db[key]
    return result_dict

def shelve_op_save(pairs):
    with shelve.open("config") as db:
        for key, value in pairs.items():
            db[key] = value



def con_timeout(builder):
    if builder.connection_timeout:
        builder.root.after_cancel(builder.connection_timeout)
    builder.connection_timeout = builder.root.after(1800000, close_pool)

def test_exist(path):
    if os.path.exists(path):
        print(path)
    else:
        print("not exits")

def pivot(data):
    flow, process, on_hold = zip(*[(rec[2], (rec[6], rec[3]), rec[7]) for rec in data])
    agg_flow = Counter(flow)
    agg_process = Counter(process)

    '''process_key = agg_process.keys()
    number_on_hold = tuple(record for record in data for k in process_key if record[3]==k and record[7]=="保留中")
    print(number_on_hold)'''

    agg_on_hold = Counter(on_hold)
    total = {"Total":len(data)}
    return (agg_flow, agg_process, agg_on_hold, total)

def dict_addup(dict, keyname, number):
    dict[keyname] = dict.get(keyname, 0) + number

def table_sort(data : list | tuple, index : int, desc : bool) -> list:
    sorted_data = list(sorted(data, key=lambda item:(0, int(item[index])) if item[index].isdigit() else (1, item[index]), reverse=desc))
    return sorted_data

class query_:
    def __init__(self, key: List[str], mode: str, plant_str: str):
        self.key = key
        self.mode = mode
        self.plant = plant_str

    def copy(self, key: List[str] = None, mode: str = None, plant_str: str = None) -> Self:        
        new_instance = query_(
            key= key if key else self.key,
            mode= mode if mode else self.mode,
            plant_str= plant_str if plant_str else self.plant
        )
        return new_instance


def validate_text(textdata: str, current_mode_str: str, plant: str, dup_list: List[str] = None) -> query_ | Dict[str, query_]:
    
    input_text = textdata
    if not input_text:
        raise E.NoKeyExc()

    split_mode = 'lot' if current_mode_str in 'itvshy' else 'like' if current_mode_str == 'a' else 'prodcode'

    #print(f'input = {input_text} split_mode = {split_mode}')
    extracted = regex_split(input_text, split_mode)

    if current_mode_str in 'fp' and len(extracted) > 1:
        raise E.ExtraKeyExc()

    
    if not extracted:
        raise E.NoKeyExc()
    if current_mode_str == 'a':
        holder = {key : query_([key], 'a', plant) for key in extracted if not key in dup_list}
        return holder
    
    
    
    query_obj = query_(extracted, current_mode_str, plant)
    
    return query_obj


def validate_return(records: List[tuple], mode: str) -> Dict[str, pd.DataFrame]:
    if not records:
       raise E.NoDataExc()
    
    df_dict = {}
    if mode not in 'tpf':
        ACS = generate_col_settings(mode, True)
        field_plain = [h['field'] for h in ACS]
        try:
            records_to_df = pd.DataFrame.from_records(records, columns=field_plain)
        except Exception as e:
            raise E.InternalErrorExc(f'Error converting records to DataFrame: {str(e)}')
        if mode in 'iew':
            records_to_df['Selected'] = ['-'] * len(records_to_df)
        if mode in ('we'):
            grouped_df = {
                db : group for db, group in records_to_df.groupby('productcode')
            }
            return grouped_df
        df_dict[mode] = records_to_df
        return df_dict
    df_dict[mode] = records
    return df_dict

class Relation_Tree:
    def __init__(self, id: str, description: str = None, children: Self = None, parent: Self = None):
        self.id = id
        self.description = description
        self.level = 1 if parent is None else parent.level + 1
        self.children = children
        self.parent = parent
        

    def __repr__(self):
        return self.render()
    
    def render(self) -> List[Tuple[str | int]]:
        return {
            'id' : self.id,
            'description' : self.description,
            'level' : self.level,
            'type' : 'ろう付けブランク' if 'B' in self.description[2:] else 'EBブランク' if 'W' in self.description else '2次加工中間品' if '-2' in self.description else '3次加工中間品' if '-3' in self.description else '製品' if not '-' in self.description else '中間品',
            'children' : [self.children.render()] if self.children else []
            
        }

    def create_child(self, id: str, description: str = None, children: Self = None) -> Self:
        new_tree = Relation_Tree(id, description, children, self)
        self.children = new_tree
        return new_tree

    def find_level(self, level: int) -> Self:
        if self.children:
            if self.children.level == level - 1:
                    return self.children
            else:
                return self.children.find_level(level)
        else:
            return self
    
    


def tree_parser(tree_cluster: List[Tuple[str | int | None]]):

    tree_list: Dict[str, Relation_Tree] = {}

    for node in tree_cluster:
        if node[4] == 1:
            new_tree = Relation_Tree(node[0], node[3])
            tree_list[node[0]] = tree_list.get(node[0], new_tree)
        if node[4] > 1:
            parent_node = node[6]
            parent_tree = tree_list[parent_node]
            parent_level = parent_tree.find_level(node[4])
            parent_level.create_child(node[0], node[3])
    
    return tuple(tree.render() for tree in list(tree_list.values()))

def search_tree(tree_cluster: List[ui.tree], db: ui.input = None, search_type: ui.input = None):
    db = db.value.upper()
    search_type = search_type.value
    def search_node(tree: List[dict] | Dict[str, str | dict], db: str = None, search_type: str = None) -> List[str] | None:
        if not db and not search_type:
            return None
        search_field_db = 'description'
        search_field_type = 'type'
        result = []
        if tree[search_field_db] == db or tree[search_field_type] == search_type:
            result.append(tree['id'])

        child = tree.get('children', None)
        if child:
            child = child[0]
            result += search_node(child, db, search_type)
        return result
    
    for t_obj in tree_cluster:
        result = []
        t = t_obj.props['nodes'][0]
        result = search_node(t, db, search_type)
        if result:
            t_obj.tick(result)
    
    return result


def html_parser(html: str) -> str:
    result = HTMLParser.feed(html)
    return result

def save_tag_on_server_single(data_list: List[Dict[str, str]], key: str, value: str | int):
    for item in data_list:
        if item['lotid'] == key:
            item.__setitem__('Selected', value)

def save_tag_on_server_batch(data_list: List[Dict[str, str]], keys: List[str], value: str | int):
    for item in data_list:
        if item['lotid'] in keys:
            item.__setitem__('Selected', value)
            keys.remove(item['lotid'])

async def modify(lid: List[Tuple[str | int]], mode: str, **kwargs) -> bool:
    #idlength = len(lid) + 1
    #placeholder_id = ', '.join(f':{i}' for i in range(1, idlength))
    datas = kwargs.get('datas', {})
    obj = LotObject(lid='', cmd=Odbconn_Cursor.MODE_BOK[mode][0], att=Odbconn_Cursor.MODE_BOK[mode][1], cond=kwargs.get("cond",(0,0)), plant=kwargs.get("plant","5501"))
    sqlstring = SD(obj)
    #print(sqlstring, lid)
    result = await OC.sql_plain_modify({sqlstring: lid})
    return result


'''
def automatedrefresh(boolvar, funcs, mins):
    for func in funcs.keys():
        func_obj, func_para = funcs[func][0], funcs[func][1]
        EV(mins).minutes.do(func_obj, func_para)
    while boolvar():
        run_pending()
        sleep(5)

def re_automatedrefresh(boolvar, funcs, mins):
    period = mins * 600
    for func in funcs.keys():
        func_obj, func_para = funcs[func][0], funcs[func][1]
        EV(mins).minutes.do(func_obj, func_para)
    while boolvar():
        run_pending()
        sleep(5)

def kill_after(seconds, window):
    sleep(seconds)
    window.destroy()
'''
