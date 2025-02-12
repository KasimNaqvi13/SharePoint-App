codeunit 99990 "Sharepoint Management"
{
    Permissions = tabledata "Dimension Set Entry" = RIMD;
    procedure SaveFile(FileDirectory: Text; FileName: Text; IS: InStream; RecordID: RecordId; Inst: InStream): Boolean
    var
        SharePointFile: Record "SharePoint File" temporary;
        IsSuccess: Boolean;
        Diag: Interface "HTTP Diagnostics";
        SharepointFolder: Record "SharePoint Folder" temporary;
        aadTenantId: Text;
        E2ESharePoint: Record "SharePoint Lists";
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
            SearchString := CopyStr(FileNameWithoutExtension, StartPos, (EndPos + 1) - StartPos)
        else
            SearchString := FileNameWithoutExtension;

        E2ESharePoint.Reset();
        E2ESharePoint.SetRange("Record ID", RecordID);
        E2ESharePoint.SetFilter(Name, SearchString + '*' + FileExtension);
        if not E2ESharePoint.FindFirst() then begin
            SharepointSetup.Get();
            InitializeConnection();
            AadTenantId := GetAadTenantNameFromBaseUrl(SharepointSetup."Initialize URL");
            SharePointClient.Initialize(SharepointSetup."Initialize URL", GetSharePointAuthorization(AadTenantId));
            CreateFolders(RecordID, FileDirectory, SharepointFolder);// test
            if SharePointClient.AddFileToFolder(PathUri, FileName, IS, SharePointFile) then begin
                IsSuccess := true;
                OnAfterAddFileToFoldeerSucess(SharepointFolder."Server Relative Url", FileName, IS, SharePointFile, RecordID, IsSuccess, FileDirectory, FileName);
            end
        end
        else begin
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
                        E2ESharePoint.Reset();
                        E2ESharePoint.SetRange("Record ID", RecordID);
                        E2ESharePoint.SetRange(Name, FileName);
                        if E2ESharePoint.FindFirst() then begin
                            if E2ESharePoint.Get(E2ESharepoint."Unique Id", E2ESharePoint."Record ID", E2ESharepoint.Name) then begin
                                if DeleteFile(E2ESharepoint."Server Relative Url") then
                                    E2ESharePoint.Delete();
                            end;
                        end;
                        FileName := FileNameWithoutExtension + '.' + FileExtension;
                        if SharePointClient.AddFileToFolder(PathUri, FileName, IS, SharePointFile) then begin
                            IsSuccess := true;
                            OnAfterAddFileToFoldeerSucess(SharepointFolder."Server Relative Url", FileName, IS, SharePointFile, RecordID, IsSuccess, FileDirectory, FileName);
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
                        E2ESharePoint.Reset();
                        E2ESharePoint.SetRange("Record ID", RecordID);
                        E2ESharePoint.SetFilter(Name, SearchString + '*' + FileExtension);
                        if E2ESharePoint.FindSet() then begin
                            FileVersion := E2ESharePoint.Count + 1;
                        end;
                        FileNameWithoutExtension := FileNameWithoutExtension + '(' + Format(FileVersion) + ')';
                        FileName := FileNameWithoutExtension + '.' + FileExtension;
                        if SharePointClient.AddFileToFolder(PathUri, FileName, IS, SharePointFile) then begin
                            IsSuccess := true;
                            OnAfterAddFileToFoldeerSucess(SharepointFolder."Server Relative Url", FileName, IS, SharePointFile, RecordID, IsSuccess, FileDirectory, FileName);
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


    procedure OpenFile(FileDirectory: Text)
    var
        SharePointFolder: Record "SharePoint Folder" temporary;
        SharePointFile: Record "SharePoint File" temporary;
        FileMgt: Codeunit "File Management";
        FileName: Text;
        FileNotFoundErr: Label 'File not found.';
    begin
        InitializeConnection();

        FileName := FileMgt.GetFileName(FileDirectory);
        if FileName = '' then
            Error(FileNotFoundErr);

        FileDirectory := FileMgt.GetDirectoryName(FileDirectory);

        if not SharePointClient.GetFolderFilesByServerRelativeUrl(FileDirectory, SharePointFile) then
            Error(FileNotFoundErr);

        SharePointFile.SetRange(Name, FileName);
        if SharePointFile.FindFirst() then begin
            SharePointClient.DownloadFileContent(SharePointFile.OdataId, FileName);
        end else
            Error(FileNotFoundErr);
    end;

    procedure DeleteFile(DeleteDirectory: Text): Boolean
    var
        SharePointFolder: Record "SharePoint Folder" temporary;
        SharePointFile: Record "SharePoint File" temporary;
        FileMgt: Codeunit "File Management";
        FileName: Text;
        FileNotFoundErr: Label 'File not found.';
        SharepointSite: Text;
        FinalURL: Text;
    begin
        InitializeConnection();
        if SharepointSetup.Get() then;
        FinalURL := SharepointSetup."Sharepoint Site" + DeleteDirectory;
        FileName := FileMgt.GetFileName(FinalURL);

        DeleteDirectory := FileMgt.GetDirectoryName(DeleteDirectory);

        if not SharePointClient.GetFolderFilesByServerRelativeUrl(DeleteDirectory, SharePointFile) then
            Error(FileNotFoundErr);

        SharePointFile.SetRange(Name, FileName);
        if SharePointFile.FindFirst() then begin
            SharePointClient.DeleteFile(SharePointFile.OdataId);
            // Message('File has been deleted successfully ');
            exit(true);
        end else
            Error(FileNotFoundErr);
    end;

    procedure GetDocumentsRootFiles(var SharepointFolder: Record "SharePoint Folder" temporary; var SharepointFile: Record "SharePoint File"): Text
    var
        SharePointList: Record "SharePoint List" temporary;
    begin
        InitializeConnection();
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

    procedure GetFilesFromServerRelativeURL(ServerRelativeURL: Text; var SharepointFolder: Record "SharePoint Folder" temporary;
        var SharepointFile: Record "SharePoint File")
    begin
        InitializeConnection();
        SharePointClient.GetFolderFilesByServerRelativeUrl(ServerRelativeURL, SharePointFile);
        SharePointClient.GetSubFoldersByServerRelativeUrl(ServerRelativeURL, SharePointFolder);
    end;

    procedure InitializeConnection(): Boolean
    var
        AadTenantId: Text;
        Diag: Interface "HTTP Diagnostics";
        SharePointList: Record "SharePoint List" temporary;
    begin
        if Connected then
            exit;

        if SharepointSetup.Get() then;
        SharePointList.DeleteAll();

        AadTenantId := GetAadTenantNameFromBaseUrl(SharepointSetup."Initialize URL");
        SharePointClient.Initialize(SharepointSetup."Initialize URL", GetSharePointAuthorization(AadTenantId));
        SharePointClient.GetLists(SharePointList);
        Diag := SharePointClient.GetDiagnostics();

        if (not Diag.IsSuccessStatusCode()) then
            Error(DiagError, Diag.GetErrorMessage());

        if Diag.IsSuccessStatusCode() then
            Connected := true
        else
            if Diag.GetHttpStatusCode() = 480 then
                Connected := false;

        exit(Connected);
    end;

    procedure GetSharePointAuthorization(AadTenantId: Text): Interface "SharePoint Authorization"
    var
        SharePointAuth: Codeunit "SharePoint Auth.";
        // Scopes: List of [Text];
        Scope: Text;
        ClientSecret: SecretText;
    begin

        if SharepointSetup.Get() then;
        ClientSecret := SharepointSetup."Client Secret";
        // Scopes.Add('00000003-0000-0ff1-ce00-000000000000/.default');
        Scope := '00000003-0000-0ff1-ce00-000000000000/.default';
        exit(SharePointAuth.CreateAuthorizationCode(AadTenantId, SharepointSetup."Client ID", ClientSecret, Scope));
    end;

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

    // for mail purpose

    procedure OpenFileInstream(FileDirectory: Text; var DocumentInStream: InStream; var TempBlobCu: codeunit "Temp Blob")
    var
        SharePointFolder: Record "SharePoint Folder" temporary;
        SharePointFile: Record "SharePoint File" temporary;
        FileMgt: Codeunit "File Management";
        FileName: Text;
        FileNotFoundErr: Label 'File not found.';
        TempText: Text;
    begin
        InitializeConnection();

        FileName := FileMgt.GetFileName(FileDirectory);
        if FileName = '' then
            Error(FileNotFoundErr);

        FileDirectory := FileMgt.GetDirectoryName(FileDirectory);

        if not SharePointClient.GetFolderFilesByServerRelativeUrl(FileDirectory, SharePointFile) then
            Error(FileNotFoundErr);

        SharePointFile.SetRange(Name, FileName);
        if SharePointFile.FindFirst() then begin
            SharePointClient.DownloadFileContentByServerRelativeUrl(SharePointFile."Server Relative Url", DocumentInStream);
        end else
            Error(FileNotFoundErr);
    end;

    var
        Connected: Boolean;
        SharePointClient: Codeunit "SharePoint Client";
        DiagError: Label 'Sharepoint Management error:\\%1';

    procedure CreateFolders(RecordID: RecordId; FileDirectory: Text; SharepointFolder: Record "SharePoint Folder")
    var
        RecRef: RecordRef;
        FieldRef: FieldRef;
        keyRef: KeyRef;
        I: Integer;
        CompanyInfoRec: Record "Company Information";
        EnvironmentInfo: Codeunit "Environment Information";
    begin
        CompanyInfoRec.Get();
        Clear(PathUri);
        Clear(folderPath);
        RecRef := RecordID.GetRecord();

        if EnvironmentInfo.IsSaaS() then begin
            //Environment Name
            folderPath := FileDirectory + '/' + EnvironmentInfo.GetEnvironmentName();
            CreatedDefultFolder(folderPath, SharepointFolder);
        end
        else if EnvironmentInfo.IsOnPrem() then begin
            // On Prem
            folderPath := folderPath + '/' + 'On Prem';
            CreatedDefultFolder(folderPath, SharepointFolder);
        end;

        // User Wise
        folderPath := folderPath + '/' + UserId;
        CreatedDefultFolder(folderPath, SharepointFolder);

        // Company Name 
        folderPath := folderPath + '/' + CompanyInfoRec.Name;
        CreatedDefultFolder(folderPath, SharepointFolder);
        // Table Name
        folderPath := folderPath + '/' + RecRef.Name;
        CreatedDefultFolder(folderPath, SharepointFolder);

        for i := 1 to RecRef.KeyIndex(1).FieldCount do begin
            FieldRef := RecRef.KeyIndex(1).FieldIndex(i);
            folderPath += '/' + Format(FieldRef.Value);
            CreatedDefultFolder(folderPath, SharepointFolder);

        end;


        PathUri := folderPath;
    end;

    procedure CreatedDefultFolder(folderPath: Text; SharepointFolder: Record "SharePoint Folder")
    begin
        if SharePointClient.CreateFolder(FolderPath, SharepointFolder) then begin

        end;
    end;

    var
        PathUri: Text[250];
        SharepointSetup: Record "Sharepoint Setup";
        folderPath: Text;

    [IntegrationEvent(false, false)]
    local procedure OnAfterAddFileToFoldeerSucess("Server Relative Url": Text; FileName: Text; IS: InStream; SharePointFile: Record "SharePoint File"; RecordID: RecordId; IsSuccess: Boolean; FileDirectory: Text; FileDirectoryFileName: Text)
    begin

    end;
}