codeunit 99990 "Sharepoint Management"
{
    Permissions = tabledata "SharePoint List Log" = rimd;
    #region Connection
    procedure InitializeConnection(): Boolean
    var
        AadTenantId: Text;
        Diag: Interface "HTTP Diagnostics";
        SharePointList: Record "SharePoint List" temporary;
    begin
        if Connected then
            exit;

        SharepointSetup.Get();
        SharePointList.DeleteAll();

        AadTenantId := this.GetAadTenantNameFromBaseUrl(SharepointSetup."Initialize URL");
        SharePointClient.Initialize(SharepointSetup."Initialize URL", this.GetSharePointAuthorization(AadTenantId));
        SharePointClient.GetLists(SharePointList);
        Diag := SharePointClient.GetDiagnostics();

        if not Diag.IsSuccessStatusCode() then
            Error(DiagError, Diag.GetErrorMessage());

        if Diag.IsSuccessStatusCode() then
            Connected := true
        else
            if Diag.GetHttpStatusCode() = 480 then
                Connected := false;

        exit(Connected);
    end;
    #endregion Connection

    #region AadTenantName
    procedure GetAadTenantNameFromBaseUrl(BaseUrl: Text): Text
    var
        Uri: Codeunit Uri;
        MySiteHostSuffixTxt: Label '-my.sharepoint.com', Locked = true;
        SharePointHostSuffixTxt: Label '.sharepoint.com', Locked = true;
        OnMicrosoftTxt: Label '.onmicrosoft.com', Locked = true;
        UrlInvalidErr: Label 'The Base Url %1 does not seem to be a valid SharePoint Online Url.', Comment = '%1=BaseUrl';
        Host: Text;
    begin
        Uri.Init(BaseUrl);
        Host := Uri.GetHost();
        if not Host.EndsWith(SharePointHostSuffixTxt) then
            Error(UrlInvalidErr, BaseUrl);
        if Host.EndsWith(MySiteHostSuffixTxt) then
            exit(CopyStr(Host, 1, StrPos(Host, MySiteHostSuffixTxt) - 1) + OnMicrosoftTxt);
        exit(CopyStr(Host, 1, StrPos(Host, SharePointHostSuffixTxt) - 1) + OnMicrosoftTxt);
    end;
    #endregion AadTenantName


    #region Authorization
    procedure GetSharePointAuthorization(AadTenantId: Text): Interface "SharePoint Authorization"
    var
        SharePointAuth: Codeunit "SharePoint Auth.";
        Scope: Text;
        ClientSecret: SecretText;
    begin
        SharepointSetup.Get();
        ClientSecret := SharepointSetup."Client Secret";
        Scope := '00000003-0000-0ff1-ce00-000000000000/.default';
        exit(SharePointAuth.CreateAuthorizationCode(AadTenantId, SharepointSetup."Client ID", ClientSecret, Scope));
    end;
    #endregion Authorization


    #region Get File
    procedure OpenFile(FileDirectory: Text; var DocumentInStream: InStream; DownloadFile: Boolean)
    var
        SharePointFolder: Record "SharePoint Folder" temporary;
        SharePointFile: Record "SharePoint File" temporary;
        FileMgt: Codeunit "File Management";
        FileName: Text;
        FileNotFoundErr: Label 'File not found.';
        TempText: Text;
    begin
        this.InitializeConnection();
        FileName := FileMgt.GetFileName(FileDirectory);
        if FileName = '' then
            Error(FileNotFoundErr);

        FileDirectory := FileMgt.GetDirectoryName(FileDirectory);
        if not SharePointClient.GetFolderFilesByServerRelativeUrl(FileDirectory, SharePointFile) then
            Error(FileNotFoundErr);

        SharePointFile.SetRange(Name, FileName);
        if SharePointFile.FindFirst() then begin
            if DownloadFile then
                SharePointClient.DownloadFileContent(SharePointFile.OdataId, FileName);
            SharePointClient.DownloadFileContentByServerRelativeUrl(SharePointFile."Server Relative Url", DocumentInStream);
        end else
            Error(FileNotFoundErr);
    end;
    #endregion Get File


    #region Upload Sharepoint File
    procedure SaveFile(FileDirectory: Text; FileName: Text; IS: InStream; RecordID: RecordId): Boolean
    var
        SharePointFile: Record "SharePoint File" temporary;
        IsSuccess: Boolean;
        Diag: Interface "HTTP Diagnostics";
        SharepointFolder: Record "SharePoint Folder" temporary;
        aadTenantId: Text;
        SharePointListLog: Record "SharePoint List Log";
        FileMgt: Codeunit "File Management";
        ExsistingFileName: Text;
        ExsistingFileNameWithoutExtension: Text;
        FileNameWithoutExtension: Text;
        FileExtension: Text;
        FileVersion: Integer;
        Ost: OutStream;
        StartPos: Integer;
        EndPos: Integer;

        TempBlobCu: Codeunit "Temp Blob";
        TestInst: InStream;
        SearchString: Text;
    begin

        FileExtension := FileMgt.GetExtension(FileName);
        FileNameWithoutExtension := FileMgt.GetFileNameWithoutExtension(FileName);
        StartPos := StrPos(FileNameWithoutExtension, '(');
        EndPos := StrPos(FileNameWithoutExtension, ')');
        if (StartPos <> 0) and (EndPos <> 0) then
            SearchString := CopyStr(FileNameWithoutExtension, 1, StartPos - 1)
        else
            SearchString := FileNameWithoutExtension;

        SharePointListLog.Reset();
        SharePointListLog.SetRange("Table Record ID", RecordID);
        SharePointListLog.SetRange("Table ID", RecordID.TableNo);
        SharePointListLog.SetFilter(Name, SearchString + '*' + FileExtension);
        if not SharePointListLog.FindFirst() then begin
            SharepointSetup.Get();
            InitializeConnection();
            AadTenantId := GetAadTenantNameFromBaseUrl(SharepointSetup."Initialize URL");
            SharePointClient.Initialize(SharepointSetup."Initialize URL", GetSharePointAuthorization(AadTenantId));
            this.CreateFolders(RecordID, FileDirectory, SharepointFolder);
            if SharePointClient.AddFileToFolder(PathUri, FileName, IS, SharePointFile) then begin
                IsSuccess := true;
                SharePointListLogMgt.AddSharePointFileDetailsToRecord(SharePointFile, RecordID, IsSuccess);
            end
        end
        else begin
            //Dragon
            SharepointSetup.Get();
            InitializeConnection();
            AadTenantId := GetAadTenantNameFromBaseUrl(SharepointSetup."Initialize URL");
            SharePointClient.Initialize(SharepointSetup."Initialize URL", GetSharePointAuthorization(AadTenantId));
            CreateFolders(RecordID, FileDirectory, SharepointFolder);

            case SharepointSetup."File Upload Method" of
                "Sharepoint File Upload Method"::" ":
                    begin
                        error('File Upload Cannot be Empty');
                    end;
                "Sharepoint File Upload Method"::"Overwrite Existing File":
                    begin
                        SharePointListLog.Reset();
                        SharePointListLog.SetRange("Table Record ID", RecordID);
                        SharePointListLog.SetRange(Name, FileName);
                        if SharePointListLog.FindFirst() then begin
                            if SharePointListLog.Get(SharePointListLog."Unique Id", SharePointListLog."Table Record ID", SharePointListLog.Name) then begin
                                if DeleteFile(SharePointListLog, SharePointListLog."Server Relative Url") then
                                    SharePointListLog.Delete();
                            end;
                        end;
                        FileName := FileNameWithoutExtension + '.' + FileExtension;
                        if SharePointClient.AddFileToFolder(PathUri, FileName, IS, SharePointFile) then begin
                            IsSuccess := true;
                            SharePointListLogMgt.AddSharePointFileDetailsToRecord(SharePointFile, RecordID, IsSuccess);
                        end;
                    end;
                SharepointSetup."File Upload Method"::"Rename If Same File Name Exist":
                    begin
                        FileExtension := FileMgt.GetExtension(FileName);
                        FileNameWithoutExtension := FileMgt.GetFileNameWithoutExtension(FileName);

                        StartPos := StrPos(FileNameWithoutExtension, '(');
                        EndPos := StrPos(FileNameWithoutExtension, ')');
                        if (StartPos <> 0) and (EndPos <> 0) then
                            SearchString := CopyStr(FileNameWithoutExtension, StartPos, EndPos - StartPos)
                        else
                            SearchString := FileNameWithoutExtension;
                        SharePointListLog.Reset();
                        SharePointListLog.SetRange("Table Record ID", RecordID);
                        SharePointListLog.SetFilter(Name, SearchString + '*' + FileExtension);
                        if SharePointListLog.FindSet() then begin
                            FileVersion := SharePointListLog.Count + 1;
                        end;
                        FileNameWithoutExtension := FileNameWithoutExtension + '(' + Format(FileVersion) + ')';
                        FileName := FileNameWithoutExtension + '.' + FileExtension;
                        if SharePointClient.AddFileToFolder(PathUri, FileName, IS, SharePointFile) then begin
                            IsSuccess := true;
                            SharePointListLogMgt.AddSharePointFileDetailsToRecord(SharePointFile, RecordID, IsSuccess);
                        end;
                    end
                else begin
                    Diag := SharePointClient.GetDiagnostics();
                    if (not Diag.IsSuccessStatusCode()) then
                        Error(DiagError, Diag.GetErrorMessage());
                end;
            end;
        end;
    end;
    #endregion Upload Sharepoint File

    #region Delete Sharepoint File
    procedure DeleteFile(SharePointListLog: Record "SharePoint List Log"; DeleteDirectory: Text): Boolean
    var
        SharePointFolder: Record "SharePoint Folder" temporary;
        SharePointFile: Record "SharePoint File" temporary;
        FileMgt: Codeunit "File Management";
        FileName: Text;
        FileNotFoundErr: Label 'File not found.';
        FinalURL: Text;
        IsHandle: Boolean;
    begin
        OnBeforeDeleteSharePointFileInSite(SharePointListLog, IsHandle);
        if IsHandle then
            exit;
        this.InitializeConnection();
        SharepointSetup.Get();
        FinalURL := SharepointSetup."Sharepoint Site" + DeleteDirectory;
        FileName := FileMgt.GetFileName(FinalURL);
        DeleteDirectory := FileMgt.GetDirectoryName(DeleteDirectory);
        if not SharePointClient.GetFolderFilesByServerRelativeUrl(DeleteDirectory, SharePointFile) then
            Error(FileNotFoundErr);
        SharePointFile.SetRange(Name, FileName);
        if SharePointFile.FindFirst() then begin
            if SharePointClient.DeleteFile(SharePointFile.OdataId) then begin
                SharePointListLog.Delete(true);
                exit(true);
            end;
        end
        else
            Error(FileNotFoundErr);
    end;
    #endregion Delete Sharepoint File



    #region Root Files
    procedure GetDocumentsRootFiles(var SharepointFolder: Record "SharePoint Folder" temporary; var SharepointFile: Record "SharePoint File"): Text
    var
        SharePointList: Record "SharePoint List" temporary;
    begin
        this.InitializeConnection();
        if SharePointClient.GetLists(SharePointList) then begin
            SharePointList.SetRange(Title, 'Documents');
            if SharePointList.FindFirst() then begin
                if SharePointClient.GetDocumentLibraryRootFolder(SharePointList.OdataId, SharePointFolder) then begin
                    SharePointClient.GetFolderFilesByServerRelativeUrl(SharePointFolder."Server Relative Url", SharePointFile);
                    SharePointClient.GetSubFoldersByServerRelativeUrl(SharePointFolder."Server Relative Url", SharePointFolder);
                    exit(SharePointFolder."Server Relative Url");
                end;
            end;
        end;
    end;
    #endregion Root Files

    #region GetDetailsFromServerRelativeURL
    procedure GetFilesFromServerRelativeURL(ServerRelativeURL: Text; var SharepointFolder: Record "SharePoint Folder" temporary;
        var SharepointFile: Record "SharePoint File")
    begin
        this.InitializeConnection();
        SharePointClient.GetFolderFilesByServerRelativeUrl(ServerRelativeURL, SharePointFile);
        SharePointClient.GetSubFoldersByServerRelativeUrl(ServerRelativeURL, SharePointFolder);
    end;
    #endregion GetDetailsFromServerRelativeURL

    #region Folder Creation
    procedure CreateFolders(RecordID: RecordId; FileDirectory: Text; SharepointFolder: Record "SharePoint Folder")
    var
        RecRef: RecordRef;
        FieldRef: FieldRef;
        keyRef: KeyRef;
        I: Integer;
        CompanyInfoRec: Record "Company Information";
        EnvironmentInfo: Codeunit "Environment Information";
        folderPath: Text;
    begin
        Clear(PathUri);
        CompanyInfoRec.Get();

        if EnvironmentInfo.IsSaaS() then begin
            //Environment Name
            folderPath := FileDirectory + '/' + CopyStr(EnvironmentInfo.GetEnvironmentName(), 1, 25);
            this.CreatedDefultFolder(folderPath, SharepointFolder);
        end
        else if EnvironmentInfo.IsOnPrem() then begin
            // On Prem
            folderPath := folderPath + '/' + 'On Prem';
            this.CreatedDefultFolder(folderPath, SharepointFolder);
        end;

        // Company Name 
        folderPath := folderPath + '/' + CopyStr(CompanyInfoRec.Name, 1, 25);
        this.CreatedDefultFolder(folderPath, SharepointFolder);

        // User Wise
        folderPath := folderPath + '/' + CopyStr(UserId, 1, 25);
        this.CreatedDefultFolder(folderPath, SharepointFolder);


        RecRef := RecordID.GetRecord();
        // Table Name
        folderPath := folderPath + '/' + RecRef.Name;
        //Key wise
        this.CreatedDefultFolder(folderPath, SharepointFolder);
        for i := 1 to RecRef.KeyIndex(1).FieldCount do begin
            FieldRef := RecRef.KeyIndex(1).FieldIndex(i);
            folderPath += '/' + CopyStr(DelChr(Format(FieldRef.Value), '=', '/\'), 1, 25);
            this.CreatedDefultFolder(folderPath, SharepointFolder);
        end;
        PathUri := folderPath;
    end;

    local procedure CreatedDefultFolder(folderPath: Text; SharepointFolder: Record "SharePoint Folder")
    begin
        SharePointClient.CreateFolder(FolderPath, SharepointFolder);
    end;
    #endregion Folder Creation


    #region Global Variable
    var
        PathUri: Text[250];
        SharepointSetup: Record "Sharepoint Setup";
        Connected: Boolean;
        SharePointClient: Codeunit "SharePoint Client";
        DiagError: Label 'Sharepoint Management error:\\%1';

        SharePointListLogMgt: Codeunit "SharePoint List Log Management";
    #endregion Global Variable

    #region Events
    [IntegrationEvent(false, false)]
    local procedure OnBeforeDeleteSharePointFileInSite(var SharePointListLog: Record "SharePoint List Log"; IsHandle: Boolean)
    begin
    end;
    #endregion Events

}