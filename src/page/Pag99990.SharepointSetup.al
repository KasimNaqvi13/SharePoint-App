page 99990 "Sharepoint Setup"
{
    PageType = Card;
    UsageCategory = Administration;
    SourceTable = "Sharepoint Setup";
    InsertAllowed = false;
    DeleteAllowed = false;
    ApplicationArea = All;
    layout
    {
        area(Content)
        {
            group(setup)
            {
                field("Azure App Name"; Rec."Azure App Name")
                {
                }
                field("Client ID"; Rec."Client ID")
                {
                    AccessByPermission = codeunit "Sharepoint Management" = x;
                    Importance = Promoted;
                    ShowMandatory = true;
                    NotBlank = true;
                }
                field("Client Secret"; Rec."Client Secret")
                {
                    AccessByPermission = codeunit "Sharepoint Management" = x;
                    ShowMandatory = true;
                    NotBlank = true;
                    ExtendedDatatype = Masked;

                }
                field("Initialize URL"; Rec."Initialize URL")
                {
                    AccessByPermission = codeunit "Sharepoint Management" = x;
                    ShowMandatory = true;
                    NotBlank = true;
                }
                field("Sharepoint Site"; Rec."Sharepoint Site")
                {

                    ShowMandatory = true;
                    NotBlank = true;
                }
                field("File Upload Method"; Rec."File Upload Method")
                {
                    ShowMandatory = true;
                }
            }

            group(Directory)
            {
                field("Default Directory"; Rec."Default Directory")
                {
                    ShowMandatory = true;
                }
                field("Purchase Directory"; Rec."Purchase Directory")
                {
                }
                field("Sales Directory"; Rec."Sales Directory")
                {
                }
                field("Customer Directory"; Rec."Customer Directory")
                {

                }
                field("Vendor Directory"; Rec."Vendor Directory")
                {

                }
            }
        }
    }

    trigger OnOpenPage()
    begin
        Rec.Reset();
        if not Rec.Get() then begin
            Rec.Init();
            Rec.Insert(true);
        end;
    end;
}