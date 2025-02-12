pageextension 99990 "Sales Order Exts" extends "Sales Order" //42
{
    layout
    {
        addfirst(factboxes)
        {
            part("E2E-SharePoint Factbox"; "SharePoint Factbox")
            {
                ApplicationArea = all;
                Caption = 'E2E-Share Point Attachment';
                SubPageLink = "Document No" = field("No."), "Document Type" = field("Document Type"), "Table ID" = const(Database::"Sales Header");
            }
        }
    }
    trigger OnAfterGetRecord()
    begin
        CurrPage."E2E-SharePoint Factbox".Page.Setrecord(Rec.RecordId, true);
    end;
}