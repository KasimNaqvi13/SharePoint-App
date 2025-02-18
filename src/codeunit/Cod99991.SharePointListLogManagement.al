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



    #region Purchase Archive
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
        if not SharepointSetupRec."Purchase Order Flow" then
            exit;
        //Header
        SharePointListLog.SetRange("Table Record ID", PurchaseHeader.RecordId);
        SharePointListLog.SetRange("Table ID", Database::"Purchase Header");
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
                SharePointListLog.Reset();
                SharePointListLog.SetRange("Table Record ID", PurchLine.RecordId);
                SharePointListLog.SetRange("Table ID", Database::"Purchase Line");
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
    #endregion Purchase Archive


    #region PQ to PO
    // // Purchase Quote to Purchase order
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Purch.-Quote to Order", OnBeforeDeletePurchQuote, '', false, false)]
    local procedure "Purch.-Quote to Order_OnBeforeDeletePurchQuote"(var QuotePurchHeader: Record "Purchase Header"; var OrderPurchHeader: Record "Purchase Header"; var IsHandled: Boolean)
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointMgt: Codeunit "Sharepoint Management";
        DocumentInStream: InStream;
        SharepointSetupRec: Record "Sharepoint Setup";
        //
        PurchQuoteLineRec: Record "Purchase Line";
        PurchaseOrderLineRec: Record "Purchase Line";
    begin
        SharepointSetupRec.Get();
        if not SharepointSetupRec."Purchase Order Flow" then
            exit;
        //Header
        SharePointListLog.SetRange("Table ID", Database::"Purchase Header");
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
                if PurchaseOrderLineRec.Get(QuotePurchHeader."Document Type", QuotePurchHeader."No.", PurchQuoteLineRec."Line No.") then begin
                    SharePointListLog.Reset();
                    SharePointListLog.SetRange("Table ID", Database::"Purchase Line");
                    SharePointListLog.SetRange("Table Record ID", PurchQuoteLineRec.RecordId);
                    if SharePointListLog.FindSet() then begin
                        repeat
                            Clear(DocumentInStream);
                            SharepointMGT.OpenFile(SharePointListLog."Server Relative Url", DocumentInStream, false);
                            if SharepointSetupRec."Purchase Directory" <> '' then
                                SharepointMgt.SaveFile(SharepointSetupRec."Purchase Directory", SharePointListLog.Name, DocumentInStream, PurchaseOrderLineRec.RecordId)
                            else
                                SharepointMgt.SaveFile(SharepointSetupRec."Default Directory", SharePointListLog.Name, DocumentInStream, PurchaseOrderLineRec.RecordId);
                        until SharePointListLog.Next() = 0;
                    end;
                end;
            until PurchQuoteLineRec.Next() = 0;
        end;
    end;
    #endregion PQ to PO

    #region PO to Purch Inv
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Purch.-Post (Yes/No)", OnAfterPost, '', false, false)]
    local procedure "Purch.-Post (Yes/No)_OnAfterPost"(var PurchaseHeader: Record "Purchase Header")
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointMgt: Codeunit "Sharepoint Management";
        DocumentInStream: InStream;
        SharepointSetupRec: Record "Sharepoint Setup";
        PurchOrderLineRec: Record "Purchase Line";
        PurchInvHeaderRec: Record "Purch. Inv. Header";
        PurchInvLineRec: Record "Purch. Inv. Line";
    begin
        SharepointSetupRec.Get();
        if not SharepointSetupRec."Purchase Order Flow" then
            exit;

        PurchInvHeaderRec.Get(PurchaseHeader."Posting No.");
        //Header
        SharePointListLog.SetRange("Table ID", Database::"Purchase Header");
        SharePointListLog.SetRange("Table Record ID", PurchaseHeader.RecordId);
        if SharePointListLog.FindSet() then begin
            repeat
                Clear(DocumentInStream);
                SharepointMgt.OpenFile(SharePointListLog."Server Relative Url", DocumentInStream, false);
                if SharepointSetupRec."Purchase Directory" <> '' then
                    SharepointMgt.SaveFile(SharepointSetupRec."Purchase Directory", SharePointListLog.Name, DocumentInStream, PurchInvHeaderRec.RecordId)
                else
                    SharepointMgt.SaveFile(SharepointSetupRec."Default Directory", SharePointListLog.Name, DocumentInStream, PurchInvHeaderRec.RecordId)
            until SharePointListLog.Next() = 0;
        end;
        PurchOrderLineRec.SetRange("Document Type", PurchaseHeader."Document Type");
        PurchOrderLineRec.SetRange("Document No.", PurchaseHeader."No.");
        if PurchOrderLineRec.FindSet() then begin
            repeat
                PurchInvLineRec.Get(PurchInvHeaderRec."No.", PurchOrderLineRec."Line No.");
                SharePointListLog.Reset();
                SharePointListLog.SetRange("Table ID", Database::"Purchase Line");
                SharePointListLog.SetRange("Table Record ID", PurchOrderLineRec.RecordId);
                if SharePointListLog.FindSet() then begin
                    repeat
                        Clear(DocumentInStream);
                        SharepointMgt.OpenFile(SharePointListLog."Server Relative Url", DocumentInStream, false);
                        if SharepointSetupRec."Purchase Directory" <> '' then
                            SharepointMgt.SaveFile(SharepointSetupRec."Purchase Directory", SharePointListLog.Name, DocumentInStream, PurchOrderLineRec.RecordId)
                        else
                            SharepointMgt.SaveFile(SharepointSetupRec."Default Directory", SharePointListLog.Name, DocumentInStream, PurchOrderLineRec.RecordId)
                    until SharePointListLog.Next() = 0;
                end;
            until PurchOrderLineRec.Next() = 0;
        end;
    end;
    #endregion PO to Purch Inv



    #region Sales Archive
    [EventSubscriber(ObjectType::Codeunit, Codeunit::ArchiveManagement, OnAfterStoreSalesDocument, '', false, false)]
    local procedure ArchiveManagement_OnAfterStoreSalesDocument(var SalesHeader: Record "Sales Header"; var SalesHeaderArchive: Record "Sales Header Archive")
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointMgt: Codeunit "Sharepoint Management";
        DocumentInStream: InStream;
        SharepointSetupRec: Record "Sharepoint Setup";
        SalesLineArchiveRec: Record "Sales Line Archive";
        SalesLineRec: Record "Sales Line";
    begin
        SharepointSetupRec.Get();
        if not SharepointSetupRec."Sales Order Flow" then
            exit;
        //Header
        SharePointListLog.SetRange("Table Record ID", SalesHeader.RecordId);
        SharePointListLog.SetRange("Table ID", Database::"Sales Header");
        if SharePointListLog.FindSet() then begin
            repeat
                Clear(DocumentInStream);
                SharepointMGT.OpenFile(SharePointListLog."Server Relative Url", DocumentInStream, false);
                if SharepointSetupRec."Purchase Directory" <> '' then
                    SharepointMgt.SaveFile(SharepointSetupRec."Purchase Directory", SharePointListLog.Name, DocumentInStream, SalesHeaderArchive.RecordId)
                else
                    SharepointMgt.SaveFile(SharepointSetupRec."Default Directory", SharePointListLog.Name, DocumentInStream, SalesHeaderArchive.RecordId);
            until SharePointListLog.Next() = 0;
        end;
        //Line
        SalesLineArchiveRec.SetRange("Document Type", SalesHeaderArchive."Document Type");
        SalesLineArchiveRec.SetRange("Document No.", SalesHeaderArchive."No.");
        SalesLineArchiveRec.SetRange("Doc. No. Occurrence", SalesHeaderArchive."Doc. No. Occurrence");
        SalesLineArchiveRec.SetRange("Version No.", SalesHeaderArchive."Version No.");
        if SalesLineArchiveRec.FindSet() then begin
            // PurchaseOrderLine
            if SalesLineRec.Get(SalesHeader."Document Type", SalesHeader."No.", SalesLineArchiveRec."Line No.") then begin
                SharePointListLog.SetRange("Table Record ID", SalesLineRec.RecordId);
                if SharePointListLog.FindSet() then begin
                    repeat
                        Clear(DocumentInStream);
                        SharepointMGT.OpenFile(SharePointListLog."Server Relative Url", DocumentInStream, false);
                        if SharepointSetupRec."Purchase Directory" <> '' then
                            SharepointMgt.SaveFile(SharepointSetupRec."Purchase Directory", SharePointListLog.Name, DocumentInStream, SalesLineArchiveRec.RecordId)
                        else
                            SharepointMgt.SaveFile(SharepointSetupRec."Default Directory", SharePointListLog.Name, DocumentInStream, SalesLineArchiveRec.RecordId);
                    until SharePointListLog.Next() = 0;
                end;
            end;
        end;
    end;
    #endregion Sales Archive


    #region SQ to SO

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Sales-Quote to Order", OnAfterOnRun, '', false, false)]
    local procedure "Sales-Quote to Order_OnAfterOnRun"(var SalesHeader: Record "Sales Header"; var SalesOrderHeader: Record "Sales Header")
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointMgt: Codeunit "Sharepoint Management";
        DocumentInStream: InStream;
        SharepointSetupRec: Record "Sharepoint Setup";
        SalesQuoteLineRec: Record "Sales Line";
        SalesOrderLineRec: Record "Sales Line";
    begin
        SharepointSetupRec.Get();
        if not SharepointSetupRec."Sales Order Flow" then
            exit;
        //Header
        SharePointListLog.SetRange("Table Record ID", SalesHeader.RecordId);
        SharePointListLog.SetRange("Table ID", Database::"Sales Header");
        if SharePointListLog.FindSet() then begin
            repeat
                Clear(DocumentInStream);
                SharepointMGT.OpenFile(SharePointListLog."Server Relative Url", DocumentInStream, false);
                if SharepointSetupRec."Purchase Directory" <> '' then
                    SharepointMgt.SaveFile(SharepointSetupRec."Purchase Directory", SharePointListLog.Name, DocumentInStream, SalesOrderHeader.RecordId)
                else
                    SharepointMgt.SaveFile(SharepointSetupRec."Default Directory", SharePointListLog.Name, DocumentInStream, SalesOrderHeader.RecordId);
            until SharePointListLog.Next() = 0;
        end;
        //Line
        SalesOrderLineRec.SetRange("Document Type", SalesOrderHeader."Document Type");
        SalesOrderLineRec.SetRange("Document No.", SalesOrderHeader."No.");
        if SalesOrderLineRec.FindSet() then begin
            // PurchaseOrderLine
            if SalesQuoteLineRec.Get(SalesHeader."Document Type", SalesHeader."No.", SalesOrderLineRec."Line No.") then begin
                SharePointListLog.SetRange("Table Record ID", SalesQuoteLineRec.RecordId);
                if SharePointListLog.FindSet() then begin
                    repeat
                        Clear(DocumentInStream);
                        SharepointMGT.OpenFile(SharePointListLog."Server Relative Url", DocumentInStream, false);
                        if SharepointSetupRec."Purchase Directory" <> '' then
                            SharepointMgt.SaveFile(SharepointSetupRec."Purchase Directory", SharePointListLog.Name, DocumentInStream, SalesOrderLineRec.RecordId)
                        else
                            SharepointMgt.SaveFile(SharepointSetupRec."Default Directory", SharePointListLog.Name, DocumentInStream, SalesOrderLineRec.RecordId);
                    until SharePointListLog.Next() = 0;
                end;
            end;
        end;
    end;
    #endregion SQ to SO














































    #region Sharepoint-Deletion
    // //--------------Purchase----Start-------------------------//
    [EventSubscriber(ObjectType::Table, Database::"Purchase Header", OnAfterDeleteEvent, '', false, false)]
    local procedure OnAfterDeleteEventPurchaseHeader(var Rec: Record "Purchase Header")
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        SharePointListLog.SetRange("Table Record ID", Rec.RecordId);
        SharePointListLog.SetRange("Table ID", Database::"Purchase Header");
        if SharePointListLog.FindSet() then begin
            repeat
                SharepointMgt.DeleteFile(SharePointListLog, SharePointListLog."Server Relative Url");
            until SharePointListLog.Next() = 0;
        end;
    end;

    [EventSubscriber(ObjectType::Table, Database::"Purchase Line", OnAfterDeleteEvent, '', false, false)]
    local procedure OnAfterDeleteEventPurchaseLine(var Rec: Record "Purchase Line")
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        SharePointListLog.SetRange("Table Record ID", Rec.RecordId);
        SharePointListLog.SetRange("Table ID", Database::"Purchase Line");
        if SharePointListLog.FindSet() then begin
            repeat
                SharepointMgt.DeleteFile(SharePointListLog, SharePointListLog."Server Relative Url");
            until SharePointListLog.Next() = 0;
        end;
    end;
    // //--------------Purchase----end-------------------------//

    // //--------------Purchase Archive----start-------------------------//
    [EventSubscriber(ObjectType::Table, Database::"Purchase Header Archive", OnAfterDeleteEvent, '', false, false)]
    local procedure OnAfterDeleteEventPurchaseArchiveHeader(var Rec: Record "Purchase Header Archive")
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        SharePointListLog.SetRange("Table Record ID", Rec.RecordId);
        SharePointListLog.SetRange("Table ID", Database::"Purchase Header Archive");
        if SharePointListLog.FindSet() then begin
            repeat
                SharepointMgt.DeleteFile(SharePointListLog, SharePointListLog."Server Relative Url");
            until SharePointListLog.Next() = 0;
        end;
    end;

    [EventSubscriber(ObjectType::Table, Database::"Purchase Line Archive", OnAfterDeleteEvent, '', false, false)]
    local procedure OnAfterDeleteEventPurchaseArchiveLine(var Rec: Record "Purchase Line Archive")
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        SharePointListLog.SetRange("Table Record ID", Rec.RecordId);
        SharePointListLog.SetRange("Table ID", Database::"Purchase Line Archive");
        if SharePointListLog.FindSet() then begin
            repeat
                SharepointMgt.DeleteFile(SharePointListLog, SharePointListLog."Server Relative Url");
            until SharePointListLog.Next() = 0;
        end;
    end;
    // //--------------Purchase Archive----end-------------------------//


    // //--------------Sales----Start-------------------------//
    [EventSubscriber(ObjectType::Table, Database::"Sales Header", OnAfterDeleteEvent, '', false, false)]
    local procedure OnAfterDeleteEventSalesHeader(var Rec: Record "Sales Header")
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        SharePointListLog.SetRange("Table Record ID", Rec.RecordId);
        SharePointListLog.SetRange("Table ID", Database::"Sales Header");
        if SharePointListLog.FindSet() then begin
            repeat
                SharepointMgt.DeleteFile(SharePointListLog, SharePointListLog."Server Relative Url");
            until SharePointListLog.Next() = 0;
        end;
    end;

    [EventSubscriber(ObjectType::Table, Database::"Sales Line", OnAfterDeleteEvent, '', false, false)]
    local procedure OnAfterDeleteEventSalesLine(var Rec: Record "Sales Line")
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        SharePointListLog.SetRange("Table Record ID", Rec.RecordId);
        SharePointListLog.SetRange("Table ID", Database::"Sales Line");
        if SharePointListLog.FindSet() then begin
            repeat
                SharepointMgt.DeleteFile(SharePointListLog, SharePointListLog."Server Relative Url");
            until SharePointListLog.Next() = 0;
        end;
    end;
    // //--------------Sales----End-------------------------//

    // //--------------Sales Archive----Start-------------------------//
    [EventSubscriber(ObjectType::Table, Database::"Sales Header Archive", OnAfterDeleteEvent, '', false, false)]
    local procedure OnAfterDeleteEventSalesHeaderArchive(var Rec: Record "Sales Header Archive")
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        SharePointListLog.SetRange("Table Record ID", Rec.RecordId);
        SharePointListLog.SetRange("Table ID", Database::"Sales Header Archive");
        if SharePointListLog.FindSet() then begin
            repeat
                SharepointMgt.DeleteFile(SharePointListLog, SharePointListLog."Server Relative Url");
            until SharePointListLog.Next() = 0;
        end;
    end;

    [EventSubscriber(ObjectType::Table, Database::"Sales Line Archive", OnAfterDeleteEvent, '', false, false)]
    local procedure OnAfterDeleteEventSalesLineArchive(var Rec: Record "Sales Line Archive")
    var
        SharePointListLog: Record "SharePoint List Log";
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        SharePointListLog.SetRange("Table Record ID", Rec.RecordId);
        SharePointListLog.SetRange("Table ID", Database::"Sales Line Archive");
        if SharePointListLog.FindSet() then begin
            repeat
                SharepointMgt.DeleteFile(SharePointListLog, SharePointListLog."Server Relative Url");
            until SharePointListLog.Next() = 0;
        end;
    end;
    // //--------------Sales Archive----End-------------------------//
    #endregion Sharepoint-Deletion







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