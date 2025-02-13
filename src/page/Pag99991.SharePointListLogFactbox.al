page 99991 "SharePointListLog Factbox"
{
    ApplicationArea = All;
    Caption = 'SharePoint Lists';
    PageType = ListPart;
    SourceTable = "SharePoint List Log";
    RefreshOnActivate = true;
    LinksAllowed = True;
    Editable = false;
    InsertAllowed = false;
    DeleteAllowed = false;


    #region layout
    layout
    {
        area(content)
        {
            repeater(General)
            {

                field(Name; Rec.Name)
                {
                    Editable = false;
                    ApplicationArea = all;
                    trigger OnDrillDown()
                    var
                        inst: InStream;
                    begin
                        SharepointMgt.OpenFile(Rec."Server Relative Url", inst, true);
                    end;
                }
                field(Link; Rec.Link)
                {
                    Editable = false;
                    ApplicationArea = all;
                    ExtendedDatatype = URL;
                }
            }
        }
    }
    #endregion layout

    #region Action
    actions
    {
        area(Processing)
        {
            fileuploadaction("Upload Files")
            {
                AllowMultipleFiles = true;
                Caption = 'Upload Files';
                ApplicationArea = all;
                Gesture = RightSwipe;
                Visible = ActionVisible;
                Image = Add;
                trigger OnAction(Files: List of [FileUpload])
                var
                    CurrentFile: FileUpload;
                    Inst: InStream;
                    FromFileName: Text;
                    Ishandle: Boolean;
                begin
                    OnBeforeUpload(Rec, Getrecord, Ishandle);
                    if Ishandle then
                        exit;
                    GetParentDirectoryFolderURL(Getrecord());
                    foreach currentFile in files do begin
                        FromFileName := CurrentFile.FileName;
                        CurrentFile.CreateInStream(Inst);
                        SharepointMgt.SaveFile(FinalURL, FromFileName, Inst, Getrecord(), Inst);
                    end;
                end;
            }
            action(Download)
            {
                ApplicationArea = All;
                Caption = 'Download File';
                Image = Download;
                Visible = ActionVisible;
                trigger OnAction()
                var
                    Instr: InStream;
                    IsHandle: Boolean;
                begin
                    OnBeforeDownload(Rec);
                    if Ishandle then
                        exit;
                    SharepointMgt.OpenFile(Rec."Server Relative Url", Instr, true);
                end;
            }
            action(DeleteFile)
            {
                ApplicationArea = All;
                Caption = 'Delete File';
                Image = Delete;
                Visible = ActionVisible;

                trigger OnAction()
                var
                    Ishandle: Boolean;
                begin
                    OnBeforeDeleteSharepointFile(Rec, Getrecord(), Ishandle);
                    if Ishandle then
                        exit;
                    SharepointMgt.DeleteFile(Rec, Rec."Server Relative Url");
                end;
            }

        }
    }
    #endregion Action

    #region Local Procedure
    local procedure GetParentDirectoryFolderURL(lRecordId: RecordId) //CIT245 #Sharepoint
    var
        SharepointSetup: Record "Sharepoint Setup";
        IsHandle: Boolean;
    begin
        OnBeforeGetParentDirectoryFolderURL(lRecordId, IsHandle, FinalURL);
        if IsHandle then
            exit;

        case lRecordId.TableNo of
            Database::Vendor:
                begin
                    SharepointSetup.Get();
                    FinalURL := SharepointSetup."Vendor Directory"; // 
                end;
            Database::"Purchase Header":
                begin
                    SharepointSetup.Get();
                    FinalURL := SharepointSetup."Purchase Directory";
                end;
            Database::"Purchase Line":
                begin
                    SharepointSetup.Get();
                    FinalURL := SharepointSetup."Purchase Directory";
                end;
            Database::Customer:
                begin
                    SharepointSetup.Get();
                    FinalURL := SharepointSetup."Customer Directory";
                end;
            Database::"Sales Header":
                begin
                    SharepointSetup.Get();
                    FinalURL := SharepointSetup."Sales Directory";
                    ;
                end;
            Database::"Sales Line":
                begin
                    SharepointSetup.Get();
                    FinalURL := SharepointSetup."Sales Directory";
                end;
            else
                SharepointSetup.Get();
                FinalURL := SharepointSetup."Default Directory";
        end
    end;
    #endregion Local Procedure



    #region Global Procedure
    procedure Setrecord(variant: Variant; Visible: Boolean)
    begin
        ActionVisible := Visible;
        if variant.IsRecordId then
            GReocrdID := variant;
    end;

    procedure Getrecord(): RecordId
    begin
        exit(GReocrdID);
    end;
    #endregion Global Procedure




    #region Global Variable
    var
        GReocrdID: RecordId;
        ActionVisible: Boolean;

    #endregion Global Variable
    #region local variable
    var

        FinalURL: Text;
        ParentFolderURL: Text;
        RetrieveURL: Text;
        SharepointMgt: Codeunit "Sharepoint Management";
    #endregion local variable



    #region Events
    [IntegrationEvent(false, false)]
    local procedure OnBeforeGetParentDirectoryFolderURL(RecordId: RecordId; Ishandle: Boolean; var FinalURL: Text)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeUpload(var SharePointListLog: Record "SharePoint List Log"; RecordId: RecordId; Ishandle: Boolean)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeDownload(var SharePointListLog: Record "SharePoint List Log")
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeDeleteSharepointFile(var SharePointListLog: Record "SharePoint List Log"; RecordId: RecordId; Ishandle: Boolean)
    begin
    end;
    #endregion Events


}
