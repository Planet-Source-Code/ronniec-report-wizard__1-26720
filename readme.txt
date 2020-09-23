****************Report Wizard Info************

The idea of this application is to allow users to 
create simple reports without allowing them to get into your
application design.  Handy if people bother you for
simple report all the time like me!!

I have modified the report to look up the NWIND database but
it can realistically be used for any database.  I use an 
Oracle system which funny enough is easier to find column names
and queries.

A few things you will have to do before it works

1. Create a new query in you Nwind (Or any other) database.  I created one called
NewQry for handiness and the SQL is as follows.

SELECT Customers.CustomerID, Customers.CompanyName, Customers.ContactName, Customers.ContactTitle, Customers.Address, Customers.City, Customers.Region, Customers.PostalCode, Customers.Country, Customers.Phone, Customers.Fax, Orders.OrderID, Orders.CustomerID, Orders.EmployeeID, Orders.OrderDate, Orders.RequiredDate, Orders.ShippedDate, Orders.ShipVia, Orders.Freight, Orders.ShipName, Orders.ShipAddress, Orders.ShipCity, Orders.ShipRegion, Orders.ShipPostalCode, Orders.ShipCountry, [Order Details].OrderID, [Order Details].ProductID, [Order Details].UnitPrice, [Order Details].Quantity, [Order Details].Discount
FROM (Customers INNER JOIN Orders ON Customers.CustomerID = Orders.CustomerID) INNER JOIN [Order Details] ON Orders.OrderID = [Order Details].OrderID;


2. You then have to find the ObjectID for this new query. This is found in the MsysObjects table.
If you cannot see this then goto Tools, Options, View (Show) System Objects and check the box.
Open the MSYSOBJECTS table and go to the Name Column.  Find the name of your new query and note the ID.
This is the object id used in the sql string for the list box.  The Attribute for a column name is 6.


3.Then goto Tools/Security /User & Group Permissions and Allow users to Read Data of your MSYSQUERIES table.

As far as I can remember that is all I had to do!!

Any Q's  Drop me one!

