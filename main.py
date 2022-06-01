# This is a sample Python script.
import xlwt
import mariadb

import config

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.


# connection parameters
# conn_params = {
#     "user": "",
#     "password": "",
#     "host": "",
#     "port": 0,
#     "database": ""
# }
conn_params = config.conn_params

# Establish a connection
connection = mariadb.connect(**conn_params)
cursor = connection.cursor()

table_column_sql = " select t.table_name,t.table_comment,column_comment,column_name,data_type,column_type " \
                   " from information_schema.TABLES t,information_schema.COLUMNS c  where t.table_schema = '" + conn_params.get(
    'database') + "' " \
                  " and t.TABLE_NAME = c.TABLE_NAME " \
                  " order by t.table_name asc;"


def TableToObject():
    """
    result = [{
        "table_name": "hap_template",
        "table_comment": "模版表",
        "column": [
            {
                "column_name": "id",
                "column_comment": "主键",
                "data_type": "varchar",
                "column_type": "varchar(128)"
            },
        ]
    }]
    """
    result = []
    # query database all tables
    cursor.execute(table_column_sql)
    # fetch all tables
    temp_name = ""
    table = {}
    for (table_name, table_comment, column_comment, column_name, data_type, column_type) in cursor:
        if temp_name != table_name:
            table = {"table_name": table_name, "table_comment": table_comment, "column": []}
            result.append(table)
            temp_name = table_name

        column = {"column_name": column_name, "column_comment": column_comment, "data_type": data_type,
                  "column_type": column_type}

        table["column"].append(column)

    return result

#
book = xlwt.Workbook()

sheet = book.add_sheet('数据资产清单')

table_title = ['序号', '业务系统', '数据表中文名称', '数据表英文名称', '计划归集月份', '备注']

column_title = ['序号', '数据项中文名称', '数据项英文名称', '数据项存储类型', '数据项长度/精度', '数据项含义', '取值说明', '其它']

output_path = "/Users/xiongus/Downloads/database.xls"

def ObjectToExcel(tables):
    # table
    table_row = 0
    table_col = 0
    for title in table_title:
        sheet.write(table_row, table_col, title)
        table_col += 1

    # fetch all tables
    for table in tables:
        table_comment = table.get("table_comment")
        table_name = table.get("table_name")
        table_row += 1

        sheet.write(table_row, 0, table_row)
        table_sheet_name = table_comment
        if table_comment is None or len(table_comment) == 0:
            table_sheet_name = table_name
        # worksheet name len 30
        if len(table_sheet_name) > 30:
            table_sheet_name = table_sheet_name[0:30]

        link = 'HYPERLINK("#' + table_sheet_name + '!B1";"' + table_comment + '")'
        formula = xlwt.Formula(link)
        sheet.write(table_row, 2, formula)
        link = 'HYPERLINK("#' + table_sheet_name + '!B1";"' + table_name + '")'
        formula = xlwt.Formula(link)
        sheet.write(table_row, 3, formula)

        table_sheet = book.add_sheet(table_sheet_name)
        column_row = 0
        column_col = 0
        for title in column_title:
            table_sheet.write(column_row, column_col, title)
            column_col += 1

        columns = table.get("column")
        for column in columns:
            column_row += 1
            table_sheet.write(column_row, 0, column_row)
            column_comment = column.get("column_comment")
            comments = column_comment.split(";")
            if len(comments) > 2:
                # 数据项中文名称;数据项含义;取值说明
                table_sheet.write(column_row, 1, comments[0])
                table_sheet.write(column_row, 5, comments[1])
                table_sheet.write(column_row, 6, comments[2])
            elif len(comments) == 2:
                # 数据项中文名称/数据项含义;取值说明
                table_sheet.write(column_row, 1, comments[0])
                table_sheet.write(column_row, 5, comments[0])
                table_sheet.write(column_row, 6, comments[1])
            else:
                # 数据项中文名称
                table_sheet.write(column_row, 1, comments[0])
            table_sheet.write(column_row, 2, column.get("column_name"))
            table_sheet.write(column_row, 3, column.get("data_type"))
            table_sheet.write(column_row, 4, column.get("column_type"))
    book.save(output_path)

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    obj = []
    try:
        obj = TableToObject()
    except Exception as e:
        print("Error:", e)
    finally:
        connection.close()
    if obj is not None and len(obj) > 1:
        ObjectToExcel(obj)
