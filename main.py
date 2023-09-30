import numbers_parser as np
from numbers_parser.document import Table,Sheet,Document, Cell
from pendulum.datetime import DateTime
from typing import Iterable
from settings import FILENAME
from typing import Callable

def main()->None:
    ''' Updates my Numbers BudgetTracker with the weekly total cost of my expenses starting on Monday'''

    numbers_doc = np.Document(FILENAME)
    budget_tracker_sheet = get_sheet(numbers_doc, "DailyTracker")
    daily_tracker_table = get_table(budget_tracker_sheet, "DailyTracker")

    details = get_column_data("Details", daily_tracker_table)
    dates = get_column_data("Date", daily_tracker_table)
    costs = get_column_data("Cost", daily_tracker_table)
    weekly_costs = [("StartOfWeek (Monday)", "WeeklyCost")]
    monthly_costs = [("Month", "MonthlyCost")]
    weekly_costs.extend(calculate_weekly_costs(dates, costs))
    monthly_costs.extend(calculate_monthly_cost(dates, costs))


    if weekly_tracker_sheet:=get_sheet(numbers_doc, "WeeklyTracker"): weekly_tracker_table = get_table(weekly_tracker_sheet, "WeeklyTracker")
    else: weekly_tracker_table = create_sheet(numbers_doc, "WeeklyTracker", return_table = True)("WeeklyTracker",len(weekly_costs),len(weekly_costs[0]))
    
    if monthly_tracker_sheet:=get_sheet(numbers_doc, "MonthlyTracker"): monthly_tracker_table = get_table(monthly_tracker_sheet, "MonthlyTracker")
    else: monthly_tracker_table = create_sheet(numbers_doc, "MonthlyTracker", return_table = True)("MonthlyTracker",len(monthly_costs),len(monthly_costs[0]))
    
    
    append_data(weekly_tracker_table, weekly_costs)
    append_data(monthly_tracker_table, monthly_costs)

    if not header_exists(weekly_tracker_table, "Date"): add_col(weekly_tracker_table, ["Date"])
    if not header_exists(weekly_tracker_table, "Highest Cost Item"): add_col(weekly_tracker_table, ["Highest Cost Item"])
    if not header_exists(weekly_tracker_table, "Cost"): add_col(weekly_tracker_table, ["Cost"])

    expensive_items = most_expensive_weekly_items(dates, costs, details)
    expensive_items.insert(0, ("Date", "Highest Cost Item", "Cost"))
    add_tupled_values(weekly_tracker_table, expensive_items)
 
    numbers_doc.save(FILENAME)

def add_tupled_values(table:Table, values:list[tuple])->None:
    '''Adds a list of tuples to a table as new columns, uses the first tuple of values in a list as the headers in the table'''
    col_indices = []
    for index, tup in enumerate(values):
        if index == 0:
            for val in tup:
                if header_exists(table, val):
                    col_indices.append(get_cell(table, val, coordinates_only = True)[1])
        else:
            for col_index, value in zip(col_indices, tup):
                table.write(index, col_index, value)
    return None


def header_exists(table:Table, header:str)->bool:
    '''Returns whether the header exists in the table'''
    return True if header in [table.cell(0,y).value for y in range(table.num_cols)] else False

def get_cell(table:Table, value, coordinates_only:bool = False)->Cell|tuple[int, int]:
    '''Returns the cell for the first instance of a given value'''
    for x in range(table.num_rows):
        for y in range(table.num_cols):
            if table.cell(x,y).value == value and coordinates_only:          
                return (x,y)
            elif table.cell(x,y).value == value and not coordinates_only:
                return table.cell(x,y)

def add_to_col(table:Table, header:str, values:list, overwrite:bool = False)->None:
    '''Adds data to the header specified, appends to the bottom of the last row of the column by default, overwrite will overwrite any pre-existing values'''
    if not header_exists(table, header): raise AttributeError("Header not found in table")

    for y in range(table.num_cols): 
        if table.cell(0,y).value == header: 
            col_index = y

    if overwrite:
        for x, value in zip(range(1, len(values)+1), values):
            table.write(x, col_index, value)
    else:
        for x, value in zip(range(table.num_rows, len(values)+table.num_rows)):
            table.write(x, col_index, value)
    return None

def add_col(table:Table, values:list)->None:
    '''Adds a list of values to a table by adding a new column to the end of the table'''
    new_col = table.num_cols
    coordinates = [(x, new_col) for x in range(len(values))]
    for coordinate,value in zip(coordinates, values): table.write(*coordinate, value)
    return None
    
def append_data(table:Table, data:list[tuple])->None:
    '''Adds the data from an iterable to the table'''
    columns = len(data[0])
    rows = len(data)
    flattened = flatten_list(data)
    coordinates = [(x,y) for x in range(rows) for y in range(columns)]

    for coordinate, value in zip(coordinates, flattened): table.write(*coordinate, value)

    return None

def flatten_list(stacked_list:Iterable[Iterable])->list:
    '''Returns a list that is the result of the inner iterables spread out'''

    flattened = []
    for inner in stacked_list:
        if not isinstance(inner, Iterable): 
            flattened.append(inner)
        else:
            for item in inner: flattened.append(item)

    return flattened

def remove_duplicates(duplicates:list)->list:
    '''Removes the duplicate dates in a list containing DateTime objects'''
    
    unique = []
    for item in duplicates: 
        if item not in unique: unique.append(item)
        
    return unique

def calculate_mondays(start:DateTime, end:DateTime)->list[DateTime]:
    '''Calculates the number of Mondays between the start date and the end date'''
    mondays = []

    while not is_monday(start):start = start.subtract(days = 1)

    while start < end:
        mondays.append(start)
        start = start.add(days = 7)

    mondays.reverse()
    return mondays

def calculate_weekly_costs(dates:list[DateTime], costs:list[float|int])->list[tuple[DateTime,float|int]]:
    '''Determines the number of Mondays that should be present between the earliest date and the latest date 
    and calculates the sum from one Monday until the next not including the 7th day, i.e the next Monday'''
    earliest_date = dates[len(dates)-1]
    latest_date = dates[0]
    
    mondays = calculate_mondays(earliest_date, latest_date)
    return [(monday,round(calculate_weekly_cost(monday, dates, costs), ndigits = 2)) for monday in mondays]

def calculate_weekly_cost(start_date:DateTime, dates:list[DateTime], costs:list[float|int])->float|int:
    '''Will calculate the sum of the values passed in from the date entered up until 7 days after not including the 7th day'''

    if(not isinstance(start_date, DateTime)): raise TypeError("the start_date is not a DateTime object")
    DAYS_IN_WEEK = 7
    week = [start_date.add(days = n) for n in range(DAYS_IN_WEEK)]

    weekly_cost = round(sum([cost for date, cost in zip(dates, costs) if date in week]), ndigits=2)

    return weekly_cost         

def most_expensive_weekly_item(start_date:DateTime, dates:list[DateTime], costs:list[float|int], details:list[str])->tuple[DateTime,str,float|int]:
    '''Iterates through the dates, costs and details lists beginning from the start date to find the most expensive item from that week'''
    DAYS_IN_WEEK = 7
    week = [start_date.add(days = n) for n in range(DAYS_IN_WEEK)]
    weekly_costs = [cost for date, cost in zip(dates, costs) if date in week]
    weekly_details = [detail for date, detail in zip(dates, details) if date in week]
    highest_cost = max(weekly_costs)
    offset = costs.index(weekly_costs[0])
    corresponding_index = weekly_costs.index(highest_cost)
    return (dates[corresponding_index+offset], weekly_details[corresponding_index], highest_cost)

def most_expensive_weekly_items(dates:list[DateTime], costs:list[float|int], details:list[str])->list[tuple[DateTime, str, float|int]]:
    '''Determines the number of Mondays that should be present between the earliest date and the latest date 
    and finds the most expensive item from one Monday until the next not including the 7th day, i.e the next Monday'''
    earliest_date = dates[len(dates)-1]
    latest_date = dates[0]
    mondays = calculate_mondays(earliest_date, latest_date)
    return [most_expensive_weekly_item(monday, dates, costs, details,) for monday in mondays]


def is_monday(date:DateTime)->bool:
    '''Checks if the date is a Monday'''
    MONDAY = 0
    return date.weekday() == MONDAY         
    
def get_column_data(column_name:str, table:Table)->list:
    '''Get the data from the specified column header not including the header itself'''
    
    start = 0
    end = 0

    for index, heading in enumerate(list(table.iter_rows(max_row=1, values_only=True))[0]):
        if heading == column_name:
            start = index
            end = index+1
            break
    
    if not start and not end: raise AttributeError("The specified column name is not within the table headers")

    return [val for val in list(table.iter_cols(min_col = start, max_col = end, values_only=True))[0] if val != None and val != column_name]

def create_sheet(document:Document, sheetname:str, return_table:bool = False)-> Sheet|Callable[[str,int,int], Sheet]:
    '''Creates and returns a sheet with the specified sheetname or a callable if return_table is set to True'''
    if not return_table: return document.add_sheet(sheetname)
    return lambda tablename, num_rows, num_cols: get_table(document.add_sheet(sheetname, tablename, num_rows=num_rows, num_cols=num_cols), tablename)

def get_sheet(document:Document, sheetname:str)->Sheet|None:
    '''Gets the sheet with the specified sheet name'''
    for sheet in document.sheets:
        if sheet.name == sheetname: return sheet
    print(f"{sheetname} is not in the Numbers document")
    return None

def create_table(sheet:Sheet, tablename:str, num_rows = 0, num_cols = 0)->Table:
    '''Creates and returns a sheet with the specified sheetname if does not exist'''

    return sheet.add_table(tablename, num_cols=num_cols, num_rows=num_rows) if num_rows and num_cols else sheet.add_table(tablename)

def get_table(sheet:Sheet, tablename:str)->Table|None:
    '''Returns the table from the sheets with the specified name or returns None'''
    for table in sheet.tables:
        if table.name == tablename: return table
    
    print("Table not found in sheet")
    return None

def calculate_monthly_cost(dates:list[DateTime], costs:list[float|int])->list[tuple]:
    '''Returns a dictionary with the months being the keys and the monthly cost being the values'''
    
    months:dict[str,list[float|int]] = {}

    for date, cost in zip(dates, costs):
        month_name = to_month_name(date.month)
        month_costs = months.get(month_name, False)
        if month_costs:
            month_costs.append(cost)
        else:
            months[month_name] = [cost]
    
    months = [(month,round(sum(costs), ndigits = 2)) for month, costs in months.items()]
    
    return months

def to_month_name(month:int)->str:
    months = {
        1:"January",
        2:"February",
        3:"March",
        4:"April",
        5:"May",
        6:"June",
        7:"July",
        8:"August",
        9:"September",
        10:"October",
        11:"November",
        12:"December",
    }

    return months.get(month, f"Error {month} does not exist")

if __name__ == "__main__":
    main()
    print("Updated the BudgetTracker Numbers file")