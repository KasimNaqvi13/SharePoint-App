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
        }
        field(2; "Client ID"; Text[250])
        {
            DataClassification = EndUserIdentifiableInformation;
        }
        field(3; "Client Secret"; Text[250])
        {
            DataClassification = EndUserIdentifiableInformation;
        }
        field(4; "Initialize URL"; Text[250])
        {
            DataClassification = ToBeClassified;
        }
        field(5; "Sharepoint Site"; Text[250])
        {
            DataClassification = ToBeClassified;
        }
        // Directory
        field(6; "Default Directory"; Text[500])
        {
            DataClassification = ToBeClassified;
        }
        field(7; "Purchase Directory"; Text[500])
        {
            DataClassification = ToBeClassified;
        }
        field(8; "Sales Directory"; Text[500])
        {
            DataClassification = ToBeClassified;
        }
        field(9; "Vendor Directory"; Text[500])
        {
            DataClassification = ToBeClassified;
        }
        field(10; "Customer Directory"; Text[500])
        {
            DataClassification = ToBeClassified;
        }
        field(11; "Azure App Name"; Text[2048])
        {

        }
        field(12; "File Upload Method"; Enum "Sharepoint File Upload Method")
        {
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