table 99991 "SharePoint Lists"
{
    Caption = 'E2E Share-Point';
    DataClassification = ToBeClassified;

    fields
    {

        field(1; "Unique Id"; Guid)
        {
            Caption = 'Unique Id';
        }

        field(2; Name; Text[250])
        {
            Caption = 'Name';
        }

        field(3; Created; DateTime)
        {
            Caption = 'Created';
        }

        field(4; Length; Integer)
        {
            Caption = 'Length';
        }

        field(5; Exists; Boolean)
        {
            Caption = 'Exists';
        }

        field(6; "Server Relative Url"; Text[2048])
        {
            Caption = 'Server Relative Url';
        }

        field(7; Title; Text[250])
        {
            Caption = 'Title';
        }

        field(8; OdataId; Text[2048])
        {
            Caption = 'Odata.Id';
        }

        field(9; OdataType; Text[2048])
        {
            Caption = 'Odata.Type';
        }

        field(10; OdataEditLink; Text[2048])
        {
            Caption = 'Odata.EditLink';
        }

        field(11; Id; Integer)
        {
            Caption = 'Id';
        }
        field(12; "Record ID" ; RecordId)
        {

        }
        field(13; IsSuccess; Boolean)
        {
            Caption = 'IsSuccess';
        }
        field(14; Link ; Text[2048])
        {
            
        }
        field(15; "Document No"; Code [20] )
        {

        }
        field(16; "Document Type"; Enum "SharePoint Document Type")
        {

        }
        field(17; "Line No"; Integer)
        {

        }
        //Option Field Change As you Want 
        field(18; "Project No"; Code[20])
        {

        }
        field(19; "Project Task No"; Code[20])
        {
            
        }
        field(20; Description; Text[500])
        {

        }
        field(21; "Table ID"; Integer)
        {

        }
        field(22; "Table Name"; Text[250])
        {

        }
        field(23; "Archived Version"; Integer)
        {

        }
        field(24; "Custom Type 1"; Code[20])
        {
        }
        field(25 ; "Custom Type 2"; Code[20])
        {
        }
    }

    // Key
    keys
    {
        key(PK; "Unique Id","Record ID",Name)
        {
            Clustered = true;
        }
        key(Key2; "Document No", "Line No", Link, "Unique Id", Name)
        {
            Enabled = true;
        }

    }

}
 
