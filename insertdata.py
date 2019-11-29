import pyodbc 
import xlrd
import re

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=localhost\SQLEXPRESS;'
                      'Database=Messer;'
                      'Trusted_Connection=yes;')
#Open the spreadsheet Clientes
workbook = xlrd.open_workbook('dados.xlsx')
sheetClient=workbook.sheet_by_index(0)

for i in range(1,sheetClient.nrows):
    #Get the values from Client's Spreadsheet
    client = sheetClient.cell_value(i,0)
    city = sheetClient.cell_value(i,1)
    state= sheetClient.cell_value(i,2)
    split_name = client.split()
    first_name = split_name[0]
    last_name = split_name[1]

    #Insert into data base the values
    cursor = conn.cursor()

    #Check if the city already exists in database
    cursor.execute("SELECT * FROM Messer.dbo.City WHERE Name = '%s' AND State = '%s'"%(city,state))
    exists = cursor.fetchone()
    if exists:
        print("City data already exists in Database")
    else:
        cursor.execute('''
                        INSERT INTO Messer.dbo.City (Name, State)
                        VALUES
                        ('%s','%s')
                        '''%(city,state))
        conn.commit()
    # cursor.execute("SELECT CityID FROM Messer.dbo.City WHERE Name = '%s' AND State = '%s'"%(city,state))
    # for row in cursor.fetchall():
        city_id = cursor.CityID
    
    #Check if the Customer already exists in database
    cursor.execute("SELECT * FROM Messer.dbo.Customer WHERE FirstName = '%s' AND LastName = '%s'"%(first_name,last_name))
    exists = cursor.fetchone()
    if exists:
        print("Customer data already exists in Database")
    else:
        cursor.execute('''
                        INSERT INTO Messer.dbo.Customer (CityID,FirstName,LastName)
                        VALUES
                        ('%d','%s','%s')
                        '''%(city_id,first_name,last_name))
        conn.commit()

#Open the spreadsheet Produtos
sheetProduct=workbook.sheet_by_index(1)
for i in range(1,sheetProduct.nrows):
    name = sheetProduct.cell_value(i,0)
    price = sheetProduct.cell_value(i,1)

    cursor = conn.cursor()

    #Check if the Product already exists in database
    cursor.execute("SELECT * FROM Messer.dbo.Product WHERE Name = '%s' AND Price = '%f'"%(name,float(price)))
    exists = cursor.fetchone()
    if exists:
        print("Product data already exists in Database")
    else:
        cursor.execute('''
                        INSERT INTO Messer.dbo.Product (Name, Price)
                        VALUES
                        ('%s','%f')
                        '''%(name,float(price)))
        conn.commit()

#Open the spreadsheet Vendas
sheetSales=workbook.sheet_by_index(2)
for i in range(1,sheetSales.nrows):
    client_sales = sheetSales.cell_value(i,0)
    product= sheetSales.cell_value(i,1)
    price_sales = sheetSales.cell_value(i,2)
    quantity= sheetSales.cell_value(i,3)
    comment = sheetSales.cell_value(i,4)
    date_comment = comment[:10]
    text_comment = comment[11:]
    split_name_sale = client_sales.split()
    first_name_sale = split_name_sale[0]
    last_name_sale = split_name_sale[1]

    cursor = conn.cursor()

    cursor.execute("SELECT CustomerID FROM Messer.dbo.Customer WHERE FirstName = '%s' AND LastName = '%s'"%(first_name_sale,last_name_sale))
    #exists = cursor.fetchone()
    # if exists is None:
    #     print("Customer was not found in the Database.")
    # else:
    for row in cursor.fetchall():
        customer_id = row.CustomerID
    cursor.execute("SELECT ProductID FROM Messer.dbo.Product WHERE Name = '%s'"%(product))
    #exists = cursor.fetchone()
    # if  exists is None:
    #     print("Product was not found in the Database.")
    # else:
    for row in cursor.fetchall():
        product_id = row.ProductID             
    cursor.execute('''
                    INSERT INTO Messer.dbo.Sale (CustomerID, ProductID,Price,Amount)
                    VALUES
                    ('%d','%d','%f','%d')
                    '''%(customer_id,product_id,float(price_sales),quantity))
    conn.commit()

    # cursor.execute('''
    #                 INSERT INTO Messer.dbo.Comment (CustomerID, SaleID,Date_comment,CommentText)
    #                 ('%d','%d','%s','%s')
    #                 '''%(customer_id,product_id,date_comment,text_comment))
    # conn.commit()


#Open the spreadsheet Fatores
sheetCoeff=workbook.sheet_by_index(3)
for i in range(1,sheetCoeff.nrows):
    name_coeff = sheetCoeff.cell_value(i,0)
    percentage = sheetCoeff.cell_value(i,1)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Messer.dbo.Factor WHERE Name = '%s'"%(name_coeff))
    exists = cursor.fetchone()
    if exists:
        print("Factor data already exists in Database")
    else:
        cursor.execute('''
                        INSERT INTO Messer.dbo.Factor (Name, Percentage)
                        VALUES
                        ('%s','%f')
                        '''%(name_coeff,float(percentage)))
        conn.commit()

#Close the Data Base Connection
conn.close()