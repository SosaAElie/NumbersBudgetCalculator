import numbers_parser as np
from numbers_parser.document import Table,Sheet,Document
from pendulum.datetime import DateTime
from typing import Iterable
from settings import FILENAME

def main()->None:
    ''' Updates my Numbers BudgetTracker with the weekly total cost of my expenses starting on Monday'''

    numbers_doc = np.Document(FILENAME)
    budget_tracker_sheet = create_get_sheet(numbers_doc, "DailyTracker")
    daily_tracker_table = create_get_table(budget_tracker_sheet, "DailyTracker")

    dates = get_column_data("Date", daily_tracker_table)
    costs = get_column_data("Cost", daily_tracker_table)
    mondays = remove_duplicates([date for date in dates if is_monday(date)])
    weekly_costs = [("StartOfWeek (Monday)", "WeeklyCost")]
    weekly_costs.extend([(monday,calculate_weekly_cost(monday, dates, costs)) for monday in mondays])
    weekly_tracker_sheet = create_get_sheet(numbers_doc, "WeeklyTracker", "WeeklyTracker")
    weekly_tracker_table = create_get_table(weekly_tracker_sheet, "WeeklyTracker", num_rows=1, num_cols=2)
    append_data(weekly_tracker_table, weekly_costs)
    numbers_doc.save(FILENAME)
    
def append_data(table:Table, data:Iterable[Iterable])->None:
    '''Adds the data from an iterable to the table'''
    columns = len(data)
    rows = len(data[0])
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

def remove_duplicates(dates:list)->list:
    '''Removes the duplicate dates in a list containing DateTime objects'''
    
    unique = []
    for date in dates: 
        if date not in unique: unique.append(date)
        
    return unique

def calculate_weekly_cost(start_date:DateTime, dates:list[DateTime], costs:list[float])->float:
    '''Will calculate the sum of the values passed in from the date entered up until 7 days after not including the 7th day'''

    if(not isinstance(start_date, DateTime)): raise TypeError("the start_date is not a DateTime object")
    DAYS_IN_WEEK = 7
    week = [start_date.add(days = n) for n in range(DAYS_IN_WEEK)]

    weekly_cost = sum([cost for date,cost in zip(dates, costs) if date in week])

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

def create_get_sheet(document:Document, sheetname:str, tablename = None)-> Sheet:
    '''Creates and returns a sheet with the specified sheetname if does not exists, if it does exist it returns it only'''
   
    for sheet in document.sheets:
        if sheet.name == sheetname: return sheet

    return document.add_sheet(sheetname, tablename) if tablename else document.add_sheet(sheetname)

def create_get_table(sheet:Sheet, tablename:str, num_rows = 0, num_cols = 0)->Table:
    '''Creates and returns a sheet with the specified sheetname if does not exists, if it does exist it returns it only'''
    for table in sheet.tables:
        if table.name == tablename: return table

    return sheet.add_table(tablename, num_cols=num_cols, num_rows=num_rows) if num_rows and num_cols else sheet.add_table(tablename)

if __name__ == "__main__":
    main()