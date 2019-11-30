USE Messer
GO  
CREATE PROCEDURE PriceAdjustment
AS
UPDATE Product SET Price = Price * (SELECT SUM(Percentage) FROM Factor)
GO
EXECUTE PriceAdjustment
GO
SELECT * FROM Product

