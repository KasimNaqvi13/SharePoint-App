page 99993 "SharePoint List Logs"
{
    ApplicationArea = All;
    Caption = 'E2E Share Point List Documents';
    PageType = List;
    SourceTable = "SharePoint List Log";
    UsageCategory = Lists;
    // InsertAllowed = false;
    Editable = true;
    // ModifyAllowed = false;

    layout
    {
        area(Content)
        {
            repeater(General)
            {

                field("Unique Id"; Rec."Unique Id")
                {
                    ToolTip = 'Specifies the value of the Unique Id field.', Comment = '%';
                }
                field(Name; Rec.Name)
                {
                    ToolTip = 'Specifies the value of the Name field.', Comment = '%';
                }
                field(Created; Rec.Created)
                {
                    ToolTip = 'Specifies the value of the Created field.', Comment = '%';
                }
                field(Length; Rec.Length)
                {
                    ToolTip = 'Specifies the value of the Length field.', Comment = '%';
                }
                field(Exists; Rec.Exists)
                {
                    ToolTip = 'Specifies the value of the Exists field.', Comment = '%';
                }
                field("Server Relative Url"; Rec."Server Relative Url")
                {
                    ToolTip = 'Specifies the value of the Server Relative Url field.', Comment = '%';
                }
                field(Title; Rec.Title)
                {
                    ToolTip = 'Specifies the value of the Title field.', Comment = '%';
                }
                field(OdataId; Rec.OdataId)
                {
                    ToolTip = 'Specifies the value of the Odata.Id field.', Comment = '%';
                }
                field(OdataType; Rec.OdataType)
                {
                    ToolTip = 'Specifies the value of the Odata.Type field.', Comment = '%';
                }
                field(OdataEditLink; Rec.OdataEditLink)
                {
                    ToolTip = 'Specifies the value of the Odata.EditLink field.', Comment = '%';
                }
                field(Id; Rec.Id)
                {
                    ToolTip = 'Specifies the value of the Id field.', Comment = '%';
                }
                field("Table Record ID"; Rec."Table Record ID")
                {
                    ToolTip = 'Specifies the value of the Table Record ID field.', Comment = '%';
                }
                field(IsSuccess; Rec.IsSuccess)
                {
                    ToolTip = 'Specifies the value of the IsSuccess field.', Comment = '%';
                }
                field(Link; Rec.Link)
                {
                    ToolTip = 'Specifies the value of the Link field.', Comment = '%';
                }
                field("Table ID"; Rec."Table ID")
                {
                    ToolTip = 'Specifies the value of the Table ID field.', Comment = '%';
                }
                field("Table Name"; Rec."Table Name")
                {
                    ToolTip = 'Specifies the value of the Table Name field.', Comment = '%';
                }
                field("Archived Version"; Rec."Archived Version")
                {
                    ToolTip = 'Specifies the value of the Archived Version field.', Comment = '%';
                }
                field("Code PK 1"; Rec."Code PK 1")
                {
                    ToolTip = 'Specifies the value of the Code PK 1 field.', Comment = '%';
                }
                field("Code PK 2"; Rec."Code PK 2")
                {
                    ToolTip = 'Specifies the value of the Code PK 2 field.', Comment = '%';
                }
                field("Code PK 3"; Rec."Code PK 3")
                {
                    ToolTip = 'Specifies the value of the Code PK 3 field.', Comment = '%';
                }
                field("Integer PK 1"; Rec."Integer PK 1")
                {
                    ToolTip = 'Specifies the value of the Integer PK 1 field.', Comment = '%';
                }
                field("Integer PK 2"; Rec."Integer PK 2")
                {
                    ToolTip = 'Specifies the value of the Integer PK 2 field.', Comment = '%';
                }
                field("Integer PK 3"; Rec."Integer PK 3")
                {
                    ToolTip = 'Specifies the value of the Integer PK 3 field.', Comment = '%';
                }
                field("SharePoint Enum"; Rec."SharePoint Enum")
                {
                    ToolTip = 'Specifies the value of the SharePoint Enum field.', Comment = '%';
                }
                field("Text Pk 1"; Rec."Text Pk 1")
                {
                    ToolTip = 'Specifies the value of the Text Pk 1 field.', Comment = '%';
                }
                field("Text Pk 2"; Rec."Text Pk 2")
                {
                    ToolTip = 'Specifies the value of the Text Pk 2 field.', Comment = '%';
                }
                field("Text Pk 3"; Rec."Text Pk 3")
                {
                    ToolTip = 'Specifies the value of the Text Pk 3 field.', Comment = '%';
                }
                field("GUID PK 1"; Rec."GUID PK 1")
                {
                    ToolTip = 'Specifies the value of the GUID PK 1 field.', Comment = '%';
                }
                field("GUID PK 2"; Rec."GUID PK 2")
                {
                    ToolTip = 'Specifies the value of the GUID PK 2 field.', Comment = '%';
                }
                field("GUID PK 3"; Rec."GUID PK 3")
                {
                    ToolTip = 'Specifies the value of the GUID PK 3 field.', Comment = '%';
                }
                field("Decimal PK 1"; Rec."Decimal PK 1")
                {
                    ToolTip = 'Specifies the value of the Decimal PK 1 field.', Comment = '%';
                }
                field("Decimal PK 2"; Rec."Decimal PK 2")
                {
                    ToolTip = 'Specifies the value of the Decimal PK 2 field.', Comment = '%';
                }
                field("Decimal PK 3"; Rec."Decimal PK 3")
                {
                    ToolTip = 'Specifies the value of the Decimal PK 3 field.', Comment = '%';
                }
            }
        }
    }
}
