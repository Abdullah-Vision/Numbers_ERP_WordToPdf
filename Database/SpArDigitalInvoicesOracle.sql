-- Oracle twin of SpArDigitalInvoices (SQL Server).
-- Run this in the Oracle ERP schema that owns ARInvoices / InvItems / AppCompanies /
-- ARCustomers and the helper functions GETCONFIGNAME / GETAPPCONFIGNAME.
-- Quoted aliases keep the exact placeholder casing used in Word templates.
CREATE OR REPLACE PROCEDURE SP_AR_DIGITAL_INVOICES (
    p_id     IN NUMBER,
    p_result OUT SYS_REFCURSOR
)
AS
BEGIN
    OPEN p_result FOR
    SELECT
        aii.InvoiceNo                                              AS "InvoiceNo",
        aii.Id                                                     AS "InvoiceId",
        ai.id                                                      AS "ItemId",
        aii.InvoiceDate                                            AS "InvoiceDate",
        aii.InvoiceDueDate                                         AS "InvoiceDueDate",
        aii.FbrInvoiceNo                                           AS "FbrInvoiceNo",
        (SELECT ConfigValue FROM AppCompanyConfigs
          WHERE Module = 'Global' AND ConfigName = 'Number Path')  AS "Path",
        aii.FilePath                                               AS "FilePath",
        aii.InvoiceType                                            AS "InvoiceType",
        TO_CHAR(aii.InvoiceDate, 'YYYY-MM')                        AS "TaxPeriod",
        ac.Name                                                    AS "CustomerName",
        ac.GSTNo                                                   AS "CustomerPhNo",
        ac.Address                                                 AS "CustomerAddress",
        ac.NTNo                                                    AS "BuyerRegistrationNo",
        (SELECT Province FROM AppCities WHERE id = ac.ProvinceId)  AS "BuyerProvince",
        ccc.Name                                                   AS "CompanyName",
        ccc.AddressLine1                                           AS "CompanyAddress",
        (SELECT Province FROM AppCities WHERE id = ccc.ProvinceId) AS "SellerProvince",
        ccc.NationalTaxNumber                                      AS "CompanyNTN",
        ccc.SalesTaxNumber                                         AS "CompanySTN",
        ccc.logo                                                   AS "Companylogo",
        i.Name                                                     AS "ItemDescription",
        i.Code                                                     AS "ItemCode",
        i.PartNo                                                   AS "SRONo",
        i.AlternatePartNo                                          AS "SROItemNo",
        i.HsCode                                                   AS "HsCode",
        i.Size                                                     AS "Size",
        i.PackQty                                                  AS "PackQty",
        GETCONFIGNAME(i.Unit)                                      AS "UOM",
        ai.Qty                                                     AS "Qty",
        ai.Mrp                                                     AS "Mrp",
        ai.Rate                                                    AS "Rate",
        ai.DiscountPercentage                                      AS "DiscountPercentage",
        ai.DiscountAmount                                          AS "DiscountAmount",
        aii.IncomeTaxAmt                                           AS "IncomeTaxAmt",
        ai.SalesTaxAmount                                          AS "SaleTaxAmount",
        ai.SalesTaxPercentage                                      AS "SalesTaxPercentage",
        ai.FurtherTax                                              AS "FurtherTax",
        ai.ExtraTax                                                AS "ExtraTax",
        0                                                          AS "FED",
        0                                                          AS "WithholdingTax",
        ai.Total                                                   AS "Total",
        ai.LineTotal                                               AS "LineTotal",
        aii.InvoiceDiscount                                        AS "MasterDiscount",
        aii.GrandTotal                                             AS "GrandTotal",
        aii.ReferenceNo                                            AS "ReferenceNo",
        aii.CustomerPono                                           AS "CustomerPono",
        aii.OGPNo                                                  AS "OGPNo",
        aii.Remarks                                                AS "MasterRemarks",
        ai.Remarks                                                 AS "LineRemarks",
        GETAPPCONFIGNAME(aii.PaymentMode)                          AS "SalesType",
        GETAPPCONFIGNAME(aii.PaymentTermId)                        AS "TermsOfSale",
        (SELECT NVL(VALUE, 0) FROM AppCompanySetups
          WHERE name = 'QtyDecimal' AND ROWNUM = 1)                AS "QtyRound",
        (SELECT NVL(VALUE, 0) FROM AppCompanySetups
          WHERE name = 'RateDecimal' AND ROWNUM = 1)               AS "RateRound"
    FROM ARInvoices aii
    JOIN ARInvoiceItems ai ON ai.InvoiceId = aii.Id
    JOIN InvItems i        ON i.Id = ai.itemId
    JOIN AppCompanies ccc  ON i.CompanyId = ccc.Id
    JOIN ARCustomers ac    ON aii.CustomerId = ac.Id
    WHERE aii.Id = p_id
      AND aii.isdeleted = 0
      AND ai.isdeleted = 0;
END;
/
