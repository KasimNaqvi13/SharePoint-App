codeunit 99991 "SharePoint List Log Management"
{
    Permissions = tabledata "SharePoint List Log" = rimd;
    procedure AddSharePointFileDetailsToRecord(SharePointFile: Record "SharePoint File" temporary; RecordID: RecordId; IsSuccess: Boolean)
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointSetup: Record "Sharepoint Setup";
        RecRef: RecordRef;
        FieldRef: FieldRef;
        keyRef: KeyRef;
        I: Integer;
        CodeKeyNo: Integer;
        OptionIshandle: Boolean;
    begin

        if not IsSuccess then
            exit;

        If (SharePointFile.OdataId = '') or (SharePointFile."Server Relative Url" = '') or (SharePointFile.Name = '') then
            exit;

        SharepointSetup.Get();
        SharePointListLog.Init();
        SharePointListLog.Validate("Unique Id", SharePointFile."Unique Id");
        SharePointListLog.Validate("Table Record ID", RecordID);
        SharePointListLog.Validate(Name, SharePointFile.Name);
        SharePointListLog.Insert(true);
        SharePointListLog.Validate(Created, SharePointFile.Created);
        SharePointListLog.Validate(Length, SharePointFile.Length);
        SharePointListLog.Validate(Exists, SharePointFile.Exists);
        SharePointListLog.Validate("Server Relative Url", SharePointFile."Server Relative Url");
        SharePointListLog.Validate(Title, SharePointFile.Title);
        SharePointListLog.Validate(OdataId, SharePointFile.OdataId);
        SharePointListLog.Validate(OdataType, SharePointFile.OdataType);
        SharePointListLog.Validate(OdataEditLink, SharePointFile.OdataEditLink);
        SharePointListLog.Validate(Id, SharePointFile.Id);
        SharePointListLog.Validate("Table ID", RecordID.TableNo);
        SharePointListLog.Validate(Link, SharepointSetup."Sharepoint Site" + SharePointFile."Server Relative Url");

        if RecRef.Get(RecordID) then begin
            for I := 1 to RecRef.KeyIndex(1).FieldCount do begin
                FieldRef := RecRef.KeyIndex(1).FieldIndex(i);
                case FieldRef.Type of
                    FieldRef.Type::Code:
                        begin
                            if SharePointListLog."Code PK 1" = '' then
                                SharePointListLog.Validate("Code PK 1", FieldRef.Value)
                            else if SharePointListLog."Code PK 2" = '' then
                                SharePointListLog.Validate("Code PK 2", FieldRef.Value)
                            else if SharePointListLog."Code PK 3" = '' then
                                SharePointListLog.Validate("Code PK 3", FieldRef.Value)
                            else
                                InsertKeys(FieldRef.Type::Code, FieldRef.Value);
                        end;
                    FieldRef.Type::Integer:
                        begin
                            If SharePointListLog."Integer PK 1" = 0 then
                                SharePointListLog.Validate("Integer PK 1", FieldRef.Value)
                            else if SharePointListLog."Integer PK 2" = 0 then
                                SharePointListLog.Validate("Integer PK 2", FieldRef.Value)
                            else if SharePointListLog."Integer PK 3" = 0 then
                                SharePointListLog.Validate("Integer PK 3", FieldRef.Value)
                            else
                                InsertKeys(FieldRef.Type::Integer, FieldRef.Value);
                        end;
                    FieldRef.Type::Option:
                        begin
                            OnBeforeValidateOptionDataType(SharePointListLog, FieldRef, OptionIshandle);
                            if OptionIshandle then
                                exit;
                            SharePointListLog.Validate("SharePoint Enum", FieldRef.Value);
                        end;
                    FieldRef.Type::Text:
                        begin
                            if SharePointListLog."Text Pk 1" = '' then
                                SharePointListLog.Validate("Text Pk 1", FieldRef.Value)
                            else if SharePointListLog."Text Pk 2" = '' then
                                SharePointListLog.Validate("Text Pk 2", FieldRef.Value)
                            else if SharePointListLog."Text Pk 3" = '' then
                                SharePointListLog.Validate("Text Pk 3", FieldRef.Value)
                            else
                                InsertKeys(FieldRef.Type::Text, FieldRef.Value);
                        end;

                    FieldRef.Type::Decimal:
                        begin
                            If SharePointListLog."Decimal PK 1" = 0 then
                                SharePointListLog.Validate("Decimal PK 1", FieldRef.Value)
                            else if SharePointListLog."Decimal PK 2" = 0 then
                                SharePointListLog.Validate("Decimal PK 2", FieldRef.Value)
                            else if SharePointListLog."Decimal PK 3" = 0 then
                                SharePointListLog.Validate("Decimal PK 3", FieldRef.Value)
                            else
                                InsertKeys(FieldRef.Type::Decimal, FieldRef.Value);
                        end;

                    FieldRef.Type::Guid:
                        begin
                            if SharePointListLog."GUID PK 1" = '' then
                                SharePointListLog.Validate("GUID PK 1", FieldRef.Value)
                            else if SharePointListLog."GUID PK 2" = '' then
                                SharePointListLog.Validate("GUID PK 2", FieldRef.Value)
                            else if SharePointListLog."GUID PK 3" = '' then
                                SharePointListLog.Validate("GUID PK 3", FieldRef.Value)
                            else
                                InsertKeys(FieldRef.Type::Guid, FieldRef.Value);
                        end;
                    else
                        InsertOtherDataTypeKeys(SharePointListLog, FieldRef);
                end;
            end;
            OnBeforeModifySharePointListLog(SharePointListLog, SharePointFile, RecordID);
            SharePointListLog.Modify(true);
            OnBeforeAfterSharePointListLog(SharePointListLog, RecordID);

        end;
    end;



    #region Purchase Quote Archive
    // //Purchase Quotes ---> Purchase Quotes Archive
    // //Header
    [EventSubscriber(ObjectType::Codeunit, Codeunit::ArchiveManagement, OnAfterPurchHeaderArchiveInsert, '', false, false)]
    local procedure ArchiveManagement_OnAfterPurchHeaderArchiveInsert(var PurchaseHeaderArchive: Record "Purchase Header Archive"; PurchaseHeader: Record "Purchase Header")
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointMgt: Codeunit "Sharepoint Management";
        DocumentInStream: InStream;
        SharepointSetupRec: Record "Sharepoint Setup";
    begin
        SharepointSetupRec.Get();
        SharePointListLog.SetRange("Table ID", Database::"Purchase Header");
        SharePointListLog.SetRange("Table Record ID", PurchaseHeader.RecordId);
        if SharePointListLog.FindSet() then begin
            repeat
                Clear(DocumentInStream);
                SharepointMGT.OpenFile(SharePointListLog."Server Relative Url", DocumentInStream, false);
                if SharepointSetupRec."Purchase Directory" <> '' then
                    SharepointMgt.SaveFile(SharepointSetupRec."Purchase Directory", SharePointListLog.Name, DocumentInStream, PurchaseHeaderArchive.RecordId)
                else
                    SharepointMgt.SaveFile(SharepointSetupRec."Default Directory", SharePointListLog.Name, DocumentInStream, PurchaseHeaderArchive.RecordId);

            until SharePointListLog.Next() = 0;
        end;
    end;
    //Line
    [EventSubscriber(ObjectType::Codeunit, Codeunit::ArchiveManagement, OnAfterStorePurchLineArchive, '', false, false)]
    local procedure ArchiveManagement_OnAfterStorePurchLineArchive(var PurchHeader: Record "Purchase Header"; var PurchLine: Record "Purchase Line"; var PurchHeaderArchive: Record "Purchase Header Archive"; var PurchLineArchive: Record "Purchase Line Archive")
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointMgt: Codeunit "Sharepoint Management";
        DocumentInStream: InStream;
        SharepointSetupRec: Record "Sharepoint Setup";
    begin
        SharepointSetupRec.Get();
        SharePointListLog.SetRange("Table Record ID", PurchLine.RecordId);
        if SharePointListLog.FindSet() then begin
            repeat
                Clear(DocumentInStream);
                SharepointMGT.OpenFile(SharePointListLog."Server Relative Url", DocumentInStream, false);
                if SharepointSetupRec."Purchase Directory" <> '' then
                    SharepointMgt.SaveFile(SharepointSetupRec."Purchase Directory", SharePointListLog.Name, DocumentInStream, PurchLineArchive.RecordId)
                else
                    SharepointMgt.SaveFile(SharepointSetupRec."Default Directory", SharePointListLog.Name, DocumentInStream, PurchLineArchive.RecordId);
            until SharePointListLog.Next() = 0;
        end;
    end;
    // //////////end///////////Neeed to take the subcriber after the Header and Line insert. BC can Roll back the transaction but We Can't do that for sharepoint
    #endregion Purchase Quote Archive

    #region PQ to PO
    // // Purchase Quote to Purchase order
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Purch.-Quote to Order", OnBeforeDeletePurchQuote, '', false, false)]
    local procedure "Purch.-Quote to Order_OnBeforeDeletePurchQuote"(var QuotePurchHeader: Record "Purchase Header"; var OrderPurchHeader: Record "Purchase Header"; var IsHandled: Boolean)
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointMgt: Codeunit "Sharepoint Management";
        DocumentInStream: InStream;
        SharepointSetupRec: Record "Sharepoint Setup";
        TempBlobCu: Codeunit "Temp Blob";
        //
        PurchQuoteLineRec: Record "Purchase Line";
        PurchaseLineRec: Record "Purchase Line";
    begin
        SharepointSetupRec.Get();
        //Header
        SharePointListLog.SetRange("Table Record ID", QuotePurchHeader.RecordId);
        if SharePointListLog.FindSet() then begin
            repeat
                Clear(DocumentInStream);
                SharepointMGT.OpenFile(SharePointListLog."Server Relative Url", DocumentInStream, false);
                if SharepointSetupRec."Purchase Directory" <> '' then
                    SharepointMgt.SaveFile(SharepointSetupRec."Purchase Directory", SharePointListLog.Name, DocumentInStream, OrderPurchHeader.RecordId)
                else
                    SharepointMgt.SaveFile(SharepointSetupRec."Default Directory", SharePointListLog.Name, DocumentInStream, OrderPurchHeader.RecordId);
            until SharePointListLog.Next() = 0;
        end;
        //Line
        PurchQuoteLineRec.Reset();
        PurchQuoteLineRec.SetRange("Document No.", QuotePurchHeader."No.");
        PurchQuoteLineRec.SetRange("Document Type", QuotePurchHeader."Document Type");
        if PurchQuoteLineRec.FindSet() then begin
            repeat
                if PurchaseLineRec.Get(QuotePurchHeader."Document Type", QuotePurchHeader."No.", PurchQuoteLineRec."Line No.") then begin
                    SharePointListLog.SetRange("Table Record ID", PurchQuoteLineRec.RecordId);
                    if SharePointListLog.FindSet() then begin
                        repeat
                            Clear(DocumentInStream);
                            SharepointMGT.OpenFile(SharePointListLog."Server Relative Url", DocumentInStream, false);
                            if SharepointSetupRec."Purchase Directory" <> '' then
                                SharepointMgt.SaveFile(SharepointSetupRec."Purchase Directory", SharePointListLog.Name, DocumentInStream, PurchaseLineRec.RecordId)
                            else
                                SharepointMgt.SaveFile(SharepointSetupRec."Default Directory", SharePointListLog.Name, DocumentInStream, PurchaseLineRec.RecordId);
                        until SharePointListLog.Next() = 0;
                    end;
                end;
            until PurchQuoteLineRec.Next() = 0;
        end;
    end;
    #endregion PQ to PO



    #region Purchase Order Archive
    // Purchase Order ----> Purchase Order Archive 
    [EventSubscriber(ObjectType::Codeunit, Codeunit::ArchiveManagement, OnAfterStorePurchDocument, '', false, false)]
    local procedure ArchiveManagement_OnAfterStorePurchDocument(var PurchaseHeader: Record "Purchase Header"; var PurchaseHeaderArchive: Record "Purchase Header Archive")
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointMgt: Codeunit "Sharepoint Management";
        DocumentInStream: InStream;
        SharepointSetupRec: Record "Sharepoint Setup";
        TempBlobCu: Codeunit "Temp Blob";
        PurchaseLineArchive: Record "Purchase Line Archive";
        PurchLine: Record "Purchase Line";
    begin
        SharepointSetupRec.Get();
        //Header
        SharePointListLog.SetRange("Table Record ID", PurchaseHeader.RecordId);
        if SharePointListLog.FindSet() then begin
            repeat
                Clear(DocumentInStream);
                SharepointMGT.OpenFile(SharePointListLog."Server Relative Url", DocumentInStream, false);
                if SharepointSetupRec."Purchase Directory" <> '' then
                    SharepointMgt.SaveFile(SharepointSetupRec."Purchase Directory", SharePointListLog.Name, DocumentInStream, PurchaseHeaderArchive.RecordId)
                else
                    SharepointMgt.SaveFile(SharepointSetupRec."Default Directory", SharePointListLog.Name, DocumentInStream, PurchaseHeaderArchive.RecordId);
            until SharePointListLog.Next() = 0;
        end;
        //Line
        PurchaseLineArchive.SetRange("Document Type", PurchaseHeaderArchive."Document Type");
        PurchaseLineArchive.SetRange("Document No.", PurchaseHeaderArchive."No.");
        PurchaseLineArchive.SetRange("Doc. No. Occurrence", PurchaseHeaderArchive."Doc. No. Occurrence");
        PurchaseLineArchive.SetRange("Version No.", PurchaseHeaderArchive."Version No.");
        if PurchaseLineArchive.FindSet() then begin
            // PurchaseOrderLine
            if PurchLine.Get(PurchaseHeader."Document Type", PurchaseHeader."No.", PurchaseLineArchive."Line No.") then begin
                SharePointListLog.SetRange("Table Record ID", PurchLine.RecordId);
                if SharePointListLog.FindSet() then begin
                    repeat
                        Clear(DocumentInStream);
                        SharepointMGT.OpenFile(SharePointListLog."Server Relative Url", DocumentInStream, false);
                        if SharepointSetupRec."Purchase Directory" <> '' then
                            SharepointMgt.SaveFile(SharepointSetupRec."Purchase Directory", SharePointListLog.Name, DocumentInStream, PurchaseLineArchive.RecordId)
                        else
                            SharepointMgt.SaveFile(SharepointSetupRec."Default Directory", SharePointListLog.Name, DocumentInStream, PurchaseLineArchive.RecordId);
                    until SharePointListLog.Next() = 0;
                end;
            end;
        end;
    end;
    // Purchase Order ------Purchase Order Archive
    #endregion Purchase Order Archive

    // Dragon

    // Purchase Order ------> Posted Purchase Order Invoice
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Purch.-Post", OnFinalizePostingOnBeforeUpdateAfterPosting, '', false, false)]
    local procedure "Purch.-Post_OnFinalizePostingOnBeforeUpdateAfterPosting"(var PurchHeader: Record "Purchase Header"; var TempDropShptPostBuffer: Record "Drop Shpt. Post. Buffer" temporary; var EverythingInvoiced: Boolean; var IsHandled: Boolean; var TempPurchLine: Record "Purchase Line" temporary)
    begin
        Message('Header no %1 and Line No %2', PurchHeader."No.", TempPurchLine."Line No.");
    end;

















    // #region Sharepoint-Deletion
    // //--------------Purchase----Start-------------------------//
    // [EventSubscriber(ObjectType::Table, Database::"Purchase Header", OnAfterDeleteEvent, '', false, false)]
    // local procedure OnAfterDeleteEventPurchaseHeader(var Rec: Record "Purchase Header")
    // var
    //     E2eSharepointRec: Record "SharePoint List Log";
    //     SharepointMgt: Codeunit "Sharepoint Management";
    // begin
    //     E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
    //     if E2eSharepointRec.FindSet() then begin
    //         repeat
    //             SharepointMgt.DeleteFile(E2eSharepointRec."Server Relative Url");
    //         until E2eSharepointRec.Next() = 0;
    //     end;
    //     E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
    //     if E2eSharepointRec.FindSet() then
    //         E2eSharepointRec.DeleteAll();
    // end;

    // [EventSubscriber(ObjectType::Table, Database::"Purchase Line", OnAfterDeleteEvent, '', false, false)]
    // local procedure OnAfterDeleteEventPurchaseLine(var Rec: Record "Purchase Line")
    // var
    //     E2eSharepointRec: Record "SharePoint List Log";
    //     SharepointMgt: Codeunit "Sharepoint Management";
    // begin
    //     E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
    //     if E2eSharepointRec.FindSet() then begin
    //         repeat
    //             SharepointMgt.DeleteFile(E2eSharepointRec."Server Relative Url");
    //         until E2eSharepointRec.Next() = 0;
    //     end;
    //     E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
    //     if E2eSharepointRec.FindSet() then
    //         E2eSharepointRec.DeleteAll();
    // end;
    // //--------------Purchase----end-------------------------//

    // //--------------Purchase Archive----start-------------------------//
    // [EventSubscriber(ObjectType::Table, Database::"Purchase Header Archive", OnAfterDeleteEvent, '', false, false)]
    // local procedure OnAfterDeleteEventPurchaseArchiveHeader(var Rec: Record "Purchase Header Archive")
    // var
    //     E2eSharepointRec: Record "SharePoint List Log";
    //     SharepointMgt: Codeunit "Sharepoint Management";
    // begin
    //     E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
    //     if E2eSharepointRec.FindSet() then begin
    //         repeat
    //             SharepointMgt.DeleteFile(E2eSharepointRec."Server Relative Url");
    //         until E2eSharepointRec.Next() = 0;
    //     end;
    //     E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
    //     if E2eSharepointRec.FindSet() then
    //         E2eSharepointRec.DeleteAll();
    // end;

    // [EventSubscriber(ObjectType::Table, Database::"Purchase Line Archive", OnAfterDeleteEvent, '', false, false)]
    // local procedure OnAfterDeleteEventPurchaseArchiveLine(var Rec: Record "Purchase Line Archive")
    // var
    //     E2eSharepointRec: Record "SharePoint List Log";
    //     SharepointMgt: Codeunit "Sharepoint Management";
    // begin
    //     E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
    //     if E2eSharepointRec.FindSet() then begin
    //         repeat
    //             SharepointMgt.DeleteFile(E2eSharepointRec."Server Relative Url");
    //         until E2eSharepointRec.Next() = 0;
    //     end;
    //     E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
    //     if E2eSharepointRec.FindSet() then
    //         E2eSharepointRec.DeleteAll();
    // end;
    // //--------------Purchase Archive----end-------------------------//


    // //--------------Sales----Start-------------------------//
    // [EventSubscriber(ObjectType::Table, Database::"Sales Header", OnAfterDeleteEvent, '', false, false)]
    // local procedure OnAfterDeleteEventSalesHeader(var Rec: Record "Sales Header")
    // var
    //     E2eSharepointRec: Record "SharePoint List Log";
    //     SharepointMgt: Codeunit "Sharepoint Management";
    // begin
    //     E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
    //     if E2eSharepointRec.FindSet() then begin
    //         repeat
    //             SharepointMgt.DeleteFile(E2eSharepointRec."Server Relative Url");
    //         until E2eSharepointRec.Next() = 0;
    //     end;
    //     E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
    //     if E2eSharepointRec.FindSet() then
    //         E2eSharepointRec.DeleteAll();
    // end;

    // [EventSubscriber(ObjectType::Table, Database::"Sales Line", OnAfterDeleteEvent, '', false, false)]
    // local procedure OnAfterDeleteEventSalesLine(var Rec: Record "Sales Line")
    // var
    //     E2eSharepointRec: Record "SharePoint List Log";
    //     SharepointMgt: Codeunit "Sharepoint Management";
    // begin
    //     E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
    //     if E2eSharepointRec.FindSet() then begin
    //         repeat
    //             SharepointMgt.DeleteFile(E2eSharepointRec."Server Relative Url");
    //         until E2eSharepointRec.Next() = 0;
    //     end;
    //     E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
    //     if E2eSharepointRec.FindSet() then
    //         E2eSharepointRec.DeleteAll();
    // end;
    // //--------------Sales----End-------------------------//

    // //--------------Sales Archive----Start-------------------------//
    // [EventSubscriber(ObjectType::Table, Database::"Sales Header Archive", OnAfterDeleteEvent, '', false, false)]
    // local procedure OnAfterDeleteEventSalesHeaderArchive(var Rec: Record "Sales Header Archive")
    // var
    //     E2eSharepointRec: Record "SharePoint List Log";
    //     SharepointMgt: Codeunit "Sharepoint Management";
    // begin
    //     E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
    //     if E2eSharepointRec.FindSet() then begin
    //         repeat
    //             SharepointMgt.DeleteFile(E2eSharepointRec."Server Relative Url");
    //         until E2eSharepointRec.Next() = 0;
    //     end;
    //     E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
    //     if E2eSharepointRec.FindSet() then
    //         E2eSharepointRec.DeleteAll();
    // end;

    // [EventSubscriber(ObjectType::Table, Database::"Sales Line Archive", OnAfterDeleteEvent, '', false, false)]
    // local procedure OnAfterDeleteEventSalesLineArchive(var Rec: Record "Sales Line Archive")
    // var
    //     E2eSharepointRec: Record "SharePoint List Log";
    //     SharepointMgt: Codeunit "Sharepoint Management";
    // begin
    //     E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
    //     if E2eSharepointRec.FindSet() then begin
    //         repeat
    //             SharepointMgt.DeleteFile(E2eSharepointRec."Server Relative Url");
    //         until E2eSharepointRec.Next() = 0;
    //     end;
    //     E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
    //     if E2eSharepointRec.FindSet() then
    //         E2eSharepointRec.DeleteAll();
    // end;
    // //--------------Sales Archive----End-------------------------//
    // #endregion Sharepoint-Deletion







    #region Events
    [IntegrationEvent(false, false)]
    local procedure InsertKeys(DataType: Variant; Value: Variant)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure InsertOtherDataTypeKeys(var SharePointListLog: Record "SharePoint List Log"; FieldRef: FieldRef)
    begin
    end;


    [IntegrationEvent(false, false)]
    local procedure OnBeforeValidateOptionDataType(var SharePointListLog: Record "SharePoint List Log"; FieldRef: FieldRef; OptionIsHandle: Boolean)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeModifySharePointListLog(var SharePointListLog: Record "SharePoint List Log"; SharePointFile: Record "SharePoint File" temporary; RecordID: RecordId)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeAfterSharePointListLog(var SharePointListLog: Record "SharePoint List Log"; RecordID: RecordId)
    begin
    end;
    #endregion Events

}