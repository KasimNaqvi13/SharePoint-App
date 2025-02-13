table 99991 "SharePoint List Log"
{
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
        field(12; "Table Record ID"; RecordId)
        {
            Caption = 'Table Record ID';

        }
        field(13; IsSuccess; Boolean)
        {
            Caption = 'IsSuccess';
        }
        field(14; Link; Text[2048])
        {
            Caption = 'Link';

        }
        field(15; "Table ID"; Integer)
        {
            Caption = 'Table ID';

        }
        field(16; "Table Name"; Text[250])
        {
            Caption = 'Table Name';

        }
        field(17; "Archived Version"; Integer)
        {
            Caption = 'Archived Version';

        }

        field(18; "Code PK 1"; Code[250])
        {
            Caption = 'Code PK 1';

        }
        field(19; "Code PK 2"; Code[250])
        {
            Caption = 'Code PK 2';

        }
        field(20; "Code PK 3"; Code[250])
        {
            Caption = 'Code PK 3';

        }
        field(21; "Integer PK 1"; Integer)
        {
            Caption = 'Integer PK 1';

        }
        field(22; "Integer PK 2"; Integer)
        {
            Caption = 'Integer PK 2';

        }
        field(23; "Integer PK 3"; Integer)
        {
            Caption = 'Integer PK 3';

        }
        field(24; "SharePoint Enum"; Enum "SharePoint Document Type") // If you're Document Type not available please extends and add yours document type with enum extension 
        {
            Caption = 'SharePoint Enum';

        }
        field(25; "Text Pk 1"; Text[1048])
        {
            Caption = 'Text Pk 1';

        }
        field(26; "Text Pk 2"; Text[1048])
        {
            Caption = 'Text Pk 2';

        }
        Field(27; "Text Pk 3"; Text[1048])
        {
            Caption = 'Text Pk 3';

        }
        field(28; "GUID PK 1"; Guid)
        {
            Caption = 'GUID PK 1';

        }
        field(29; "GUID PK 2"; Guid)
        {
            Caption = 'GUID PK 2';

        }
        field(30; "GUID PK 3"; Guid)
        {
            Caption = 'GUID PK 3';

        }
        field(31; "Decimal PK 1"; Decimal)
        {
            Caption = 'Decimal PK 1';

        }
        field(32; "Decimal PK 2"; Decimal)
        {
            Caption = 'Decimal PK 2';

        }
        field(33; "Decimal PK 3"; Decimal)
        {
            Caption = 'Decimal PK 3';

        }
    }

    // Key
    keys
    {
        key(PK; "Unique Id", "Table Record ID", Name)
        {
            Clustered = true;
        }
        key(Key2; Link, "Unique Id", Name)
        {
            Enabled = true;
        }

    }

}

