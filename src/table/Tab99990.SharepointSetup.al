table 99990 "Sharepoint Setup"
{
    Access = Public;
    Caption = 'Sharepoint Setup';
    CompressionType = Page;
    ColumnStoreIndex = "Client ID";
    DataPerCompany = true;
    Description = 'Sharepoint setup table store the every confidential value and with the help of this table will store the files in sharepoint directory';
    DataCaptionFields = "Azure App Name";
    DataClassification = ToBeClassified;
    Extensible = true;
    InherentEntitlements = rimdx;
    InherentPermissions = rimdx;
    // LinkedObject = true;
#Pragma warning disable
    // LinkedInTransaction = true;
    PasteIsValid = true;
    ReplicateData = false;
    Scope = Cloud;
    TableType = Normal;
    fields
    {
        field(1; "Primary Key"; Code[10])
        {
            DataClassification = ToBeClassified;
            Caption = 'Primary Key';
        }
        field(2; "Client ID"; Text[250])
        {
            DataClassification = EndUserIdentifiableInformation;
            Caption = 'Client ID';
        }
        field(3; "Client Secret"; Text[250])
        {
            DataClassification = EndUserIdentifiableInformation;
            Caption = 'Client Secret';
        }
        field(4; "Initialize URL"; Text[250])
        {
            DataClassification = ToBeClassified;
            Caption = 'Initialize URL';
        }
        field(5; "Sharepoint Site"; Text[250])
        {
            DataClassification = ToBeClassified;
            Caption = 'Sharepoint Site';
        }
        // Directory
        field(6; "Default Directory"; Text[500])
        {
            DataClassification = ToBeClassified;
            Caption = 'Default Directory';
        }
        field(7; "Purchase Directory"; Text[500])
        {
            DataClassification = ToBeClassified;
            Caption = 'Purchase Directory';
        }
        field(8; "Sales Directory"; Text[500])
        {
            DataClassification = ToBeClassified;
            Caption = 'Sales Directory';
        }
        field(9; "Vendor Directory"; Text[500])
        {
            DataClassification = ToBeClassified;
            Caption = 'Vendor Directory';
        }
        field(10; "Customer Directory"; Text[500])
        {
            DataClassification = ToBeClassified;
            Caption = 'Customer Directory';
        }
        field(11; "Azure App Name"; Text[2048])
        {
            Caption = 'Azure App Name';

        }
        field(12; "File Upload Method"; Enum "Sharepoint File Upload Method")
        {
            DataClassification = ToBeClassified;
            Caption = 'File Upload Method';
        }
        field(13; "Purchase Order Flow"; Boolean)
        {
            Caption = 'Purchase Order Flow';
            DataClassification = ToBeClassified;

        }
        field(14; "Sales Order Flow"; Boolean)
        {
            Caption = 'Sales Order Flow';
            DataClassification = ToBeClassified;

        }


    }

    keys
    {
        key(PK; "Primary Key")
        {
            Clustered = true;
        }
    }
}