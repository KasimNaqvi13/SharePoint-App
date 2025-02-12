// page 60003 "SharePoint Subform"
// {
//     ApplicationArea = All;
//     Caption = 'E2E-SharePoint Subform';
//     PageType = List;
//     SourceTable = "SharePoint Lists";
//     Editable = false;
//     InsertAllowed = false;
//     DeleteAllowed = false;
//     UsageCategory = None;
//     RefreshOnActivate = true;


//     layout
//     {
//         area(content)
//         {
//             repeater(General)
//             {

//                 field(Name; Rec.Name)
//                 {
//                     Editable = false;
//                     ApplicationArea = all;
//                     trigger OnDrillDown()
//                     var
//                         SharepointMgt: Codeunit "Sharepoint Management";
//                     begin
//                         SharepointMgt.OpenFile(Rec."Server Relative Url");
//                     end;
//                 }
//                 field(Link; Rec.Link)
//                 {
//                     Editable = false;
//                     ApplicationArea = all;
//                     ExtendedDatatype = URL;
//                 }
//             }
//         }
//     }

//     // ------ Action--------
//     actions
//     {
//         area(Processing)
//         {
//             //Action 1
//             action(UploadFile)
//             {
//                 ApplicationArea = All;
//                 Caption = 'Upload File';
//                 Image = NewDocument;
//                 Visible = ActionVisible;
//                 trigger OnAction()
//                 var
//                     IS: InStream;
//                     TempBlob: Codeunit "Temp Blob";
//                     SharepointMgt: Codeunit "Sharepoint Management";
//                     FromFileName: Text;
//                     lRecordRef: RecordRef;
//                     Ishandle: Boolean;
//                 begin
//                     Clear(lRecordId);
//                     lRecordId := Getrecord();
//                     OnBeforeActionSharePoint(lRecordId, Ishandle);
//                     if Ishandle then
//                         exit;
//                     GetParentDirectoryFolderURL(lRecordId);
//                     IS := TempBlob.CreateInStream();
//                     UploadIntoStream('Please Upload a File', '', '', FromFileName, IS);
//                     if SharepointMgt.SaveFile(FinalURL, FromFileName, IS, lRecordId, IS) then begin
//                         Message('File has been uploaded successfully!!');
//                     end;
//                 end;

//             }
//             action(Download)
//             {
//                 ApplicationArea = All;
//                 Caption = 'Download File';
//                 Image = Download;
//                 Visible = ActionVisible;
//                 trigger OnAction()
//                 var
//                     SharepointMgt: Codeunit "Sharepoint Management";
//                 begin
//                     SharepointMgt.OpenFile(Rec."Server Relative Url");
//                 end;
//             }

//             action(DeleteFile)
//             {
//                 ApplicationArea = All;
//                 Caption = 'Delete File';
//                 Image = Delete;
//                 Visible = ActionVisible;

//                 trigger OnAction()
//                 var
//                     lRecordRef: RecordRef;
//                     lRecordId: RecordId;
//                     RecordLinks: Record "Record Link";
//                     Ishandle: Boolean;
//                     RecPurchaseHead: Record "Purchase Header";
//                 begin
//                     Clear(lRecordId);
//                     lRecordId := Getrecord();
//                     OnBeforeDeleteSharepointFile(lRecordId, Ishandle);
//                     if Ishandle then
//                         exit;
//                     DeleteFileFromSharepoint(Rec);
//                 end;
//             }

//         }

//         area(Promoted)
//         {
//             actionref(UploadFiles; UploadFile)
//             {

//             }
//             actionref(Downloads; Download)
//             {

//             }
//             actionref(DeleteFiles; DeleteFile)
//             {

//             }

//         }


//     }

//     procedure GetParentDirectoryFolderURL(lRecordId: RecordId) //CIT245 #Sharepoint
//     var
//         SharepointSetup: Record "Sharepoint Setup";
//         IsHandle: Boolean;
//     begin
//         OnBeforeGetParentDirectoryFolderURL(lRecordId, IsHandle);
//         if IsHandle then
//             exit;

//         case lRecordId.TableNo of
//             Database::"Purchase Header":
//                 begin
//                     if SharepointSetup.Get() then begin
//                         FinalURL := SharepointSetup."Purchase Directory";
//                     end;
//                 end;
//             Database::"Purchase Line":
//                 begin
//                     if SharepointSetup.Get() then begin
//                         FinalURL := SharepointSetup."Purchase Directory";
//                     end;
//                 end;
//             Database::Vendor:
//                 begin
//                     if SharepointSetup.Get() then begin
//                         FinalURL := SharepointSetup."Vendor Directory"; // 
//                     end;
//                 end;
//             Database::Customer:
//                 begin
//                     if SharepointSetup.Get() then begin
//                         FinalURL := SharepointSetup."Customer Directory";
//                     end;
//                 end;
//             Database::"Sales Header":
//                 begin
//                     if SharepointSetup.Get() then begin
//                         FinalURL := SharepointSetup."Sales Directory";
//                     end;
//                 end;
//             Database::"Sales Line":
//                 begin
//                     if SharepointSetup.Get() then begin
//                         FinalURL := SharepointSetup."Sales Directory";
//                     end;
//                 end;
//             Database::job:
//                 begin
//                     if SharepointSetup.Get() then begin
//                         FinalURL := SharepointSetup."Project Directory";
//                     end;
//                 end;
//             Database::"Job Task":
//                 begin
//                     if SharepointSetup.Get() then begin
//                         FinalURL := SharepointSetup."Project Directory";
//                     end;
//                 end;
//             else
//                 // Default Url
//                 if SharepointSetup.Get() then begin
//                     FinalURL := SharepointSetup."Default Directory";
//                 end
//         end;
//     end;

//     procedure DeleteFileFromSharepoint(E2ESharepoint: Record "SharePoint Lists")
//     var
//         SharepointMgt: Codeunit "Sharepoint Management";
//         ConfirmManagement: Codeunit "Confirm Management";
//     begin
//         if ConfirmManagement.GetResponseOrDefault('Do you want to Delete the File ?', false) then begin
//             SharepointMgt.DeleteFile(Rec."Server Relative Url");
//             if E2ESharepoint.Delete() then begin

//             end;
//         end;
//     end;


//     procedure Setrecord(variant: Variant; Visible: Boolean)
//     begin
//         ActionVisible := Visible;
//         Clear(GReocrdID);
//         if variant.IsRecordId then
//             GReocrdID := variant;
//     end;

//     procedure Getrecord(): RecordId
//     begin
//         exit(GReocrdID);
//     end;

//     var
//         GReocrdID: RecordId;

//         BooleanVar: boolean;

//     var
//         ParentFolderURL: Text;
//         FinalURL: Text;
//         RetrieveURL: Text;
//         lRecordId: RecordId;
//         ActionVisible: Boolean;


//     [IntegrationEvent(false, false)]
//     local procedure OnBeforeGetParentDirectoryFolderURL(lRecordId: RecordId; Ishandle: Boolean)
//     begin

//     end;

//     [IntegrationEvent(false, false)]
//     local procedure OnBeforeDeleteSharepointFile(lRecordId: RecordId; Ishandle: Boolean)
//     begin

//     end;

//     [IntegrationEvent(false, false)]
//     local procedure OnBeforeActionSharePoint(lRecordId: RecordId; Ishandle: Boolean)
//     begin

//     end;
// }
