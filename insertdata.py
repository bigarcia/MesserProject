import pyodbc 
import xlrd
import re

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=localhost\SQLEXPRESS;'
                      'Database=Messer;'
                      'Trusted_Connection=yes;')
#Open the spreadsheet Clientes
workbook = xlrd.open_workbook('dados.xlsx')
sheetClient=workbook.sheet_by_name(Clientes)
for i in range(sheetClient.nrows):
    #Get the values from Client's Spreadsheet
    client = sheetClient.cell_value(i,1)
    city = sheetClient.cell_value(i,2)
    state= sheetClient.cell_value(i,3)
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
    cursor.execute("SELECT CityID FROM Messer.dbo.City WHERE Name = '%s' AND State = '%s'"%(city,state))
    for row in cursor.fetchall():
        city_id = row.CityID
    
    #Check if the Customer already exists in database
    cursor.execute("SELECT * FROM Messer.dbo.Customer WHERE FirstName = '%s' AND LastName = '%s'"%(first_name,last_name))
    exists = cursor.fetchone()
    if exists:
        print("Customer data already exists in Database")
    else:
        cursor.execute('''
                        INSERT INTO Messer.dbo.Customer (CityID,FirstName,LastName)
                        VALUES
                        ('%s','%s','%s')
                        '''%(city_id,first_name,last_name))
        conn.commit()

#Open the spreadsheet Produtos
sheetProduct=workbook.sheet_by_name(Produtos)
for i in range(sheetClient.nrows):
    name = sheetProduct.cell_value(i,1)
    price = sheetProduct.cell_value(i,2)

    cursor = conn.cursor()

    #Check if the Product already exists in database
    cursor.execute("SELECT * FROM Messer.dbo.Product WHERE Name = '%s' AND Price = '%s'"%(name,price))
    exists = cursor.fetchone()
    if exists:
        print("Product data already exists in Database")
    else:
        cursor.execute('''
                        INSERT INTO Messer.dbo.Product (Name, Price)
                        VALUES
                        ('%s','%s')
                        '''%(name,price))
        conn.commit()

#Open the spreadsheet Vendas
sheetSales=workbook.sheet_by_name(Vendas)
for i in range(sheetSales.nrows):
    client_sales = sheetSales.cell_value(i,1)
    product= sheetSales.cell_value(i,2)
    price_sales = sheetSales.cell_value(i,3)
    quantity= sheetSales.cell_value(i,4)
    comment= sheetSales.cell_value(i,5)
    date_comment = re.findall(r'^\d{2}\/\d{2}\/\d{4}',comment)
    comment_text = re.split(r'^\d{2}\/\d{2}\/\d{4} ',comment)
    date_comment = date_split[0]
    text_comment = text_split[1]
    split_name_sale = client_sales.split()
    first_name_sale = split_name_sale[0]
    last_name_sale = split_name_sale[1]

    cursor = conn.cursor()

    cursor.execute("SELECT CustomerID FROM Messer.dbo.Customer WHERE FirstName = '%s' AND LastName = '%s'"%(first_name_sale,last_name_sale))
    if cursor.fetchone() is None:
        print("Customer was not found in the Database.")
    else:
        for row in cursor.fetchall():
            customer_id = row.CustomerID
        cursor.execute("SELECT ProductID FROM Messer.dbo.Product WHERE Name = '%s'"%(product))
        if cursor.fetchone() is None:
            print("Product was not found in the Database.")
        else:
            for row in cursor.fetchall():
                product_id = row.ProductID             
            cursor.execute('''
                            INSERT INTO Messer.dbo.Sale (CustomerID, ProductID,Price,Ammount)
                            VALUES
                            ('%s','%s','%s','%s')
                            '''%(customer_id,product_id,price_sales,quantity))
            conn.commit()

            cursor.execute('''
                            INSERT INTO Messer.dbo.Comment (CustomerID, SaleID,Date_comment,CommentText)
                            VALUES
                            ('%s','%s','%s','%s')
                            '''%(customer_id,product_id,date_comment,text_comment))
            conn.commit()


#Open the spreadsheet Fatores
sheetCoeff=workbook.sheet_by_name(Fatores)
for i in range(sheetCoeff.nrows):
    name_coeff = sheetCoeff.cell_value(i,1)
    percentage= sheetCoeff.cell_value(i,2)

    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Messer.dbo.Factor WHERE Name = '%s'"%(name_coeff))
    exists = cursor.fetchone()
    if exists:
        print("Factor data already exists in Database")
    else:
        cursor.execute('''
                        INSERT INTO Messer.dbo.Factor (Name, Percentage)
                        VALUES
                        ('%s','%s')
                        '''%(name_coeff,percentage))
        conn.commit()

#Close the Data Base Connection
conn.close()