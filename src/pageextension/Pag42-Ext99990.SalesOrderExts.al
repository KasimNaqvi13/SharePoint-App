pageextension 99990 "Sales Order Exts" extends "Sales Order" //42
{
    layout
    {
        addfirst(factboxes)
        {
            part("SharePointListLog Factbox"; "SharePointListLog Factbox")
            {
                ApplicationArea = all;
                Caption = 'Share Point Attachment';
                SubPageLink = "Code PK 1" = field("No."), "SharePoint Enum" = field("Document Type"), "Table ID" = const(Database::"Sales Header");
            }
        }
    }
    trigger OnAfterGetRecord()
    begin
        CurrPage."SharePointListLog Factbox".Page.Setrecord(Rec.RecordId, true);
    end;
}