import numbers_parser as np
from numbers_parser.document import Table,Sheet,Document
from pendulum.datetime import DateTime
from typing import Iterable
from settings import FILENAME
from typing import Callable

def main()->None:
    ''' Updates my Numbers BudgetTracker with the weekly total cost of my expenses starting on Monday'''

    numbers_doc = np.Document(FILENAME)
    budget_tracker_sheet = get_sheet(numbers_doc, "DailyTracker")
    daily_tracker_table = get_table(budget_tracker_sheet, "DailyTracker")

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
    numbers_doc.save(FILENAME)
    
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

    return mondays

def calculate_weekly_costs(dates:list[DateTime], costs:list[float|int])->list[tuple[DateTime,float|int]]:
    '''Determines the number of Mondays that should be present between the earliest date and the latest date 
    and calculates the sum from one Monday until the next not including the 7th day, i.e the next Monday'''
    earliest_date = dates[len(dates)-1]
    latest_date = dates[0]
    
    mondays = calculate_mondays(earliest_date, latest_date)
    data = [(monday,round(calculate_weekly_cost(monday, dates, costs), ndigits = 2)) for monday in mondays]
    data.reverse()
    return data

def calculate_weekly_cost(start_date:DateTime, dates:list[DateTime], costs:list[float|int])->float|int:
    '''Will calculate the sum of the values passed in from the date entered up until 7 days after not including the 7th day'''

    if(not isinstance(start_date, DateTime)): raise TypeError("the start_date is not a DateTime object")
    DAYS_IN_WEEK = 7
    week = [start_date.add(days = n) for n in range(DAYS_IN_WEEK)]

    weekly_cost = round(sum([cost for date, cost in zip(dates, costs) if date in week]), ndigits=2)

    return weekly_cost         

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