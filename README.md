# MeasurLink-InspPlan-Conversion

Generate inspection plans in the .QIF format that Measurlink can parse from an Excel workbook staging area.

If Measurlink tries to parse a malformed .QIF file, it will abort the transaction. This would be preferable to trying to insert and update the fields
against the production database.

Test ReadmeChange