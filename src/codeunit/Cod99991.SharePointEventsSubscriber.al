codeunit 99991 "SharePoint Events & Subscriber"
{
    Permissions = tabledata "Dimension Set Entry" = RIMD;
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Sharepoint Management", OnAfterAddFileToFoldeerSucess, '', false, false)]
    local procedure "SharepointManagementOnAfterAddFileToFoldeerSucess"("Server Relative Url": Text; FileName: Text; IS: InStream; SharePointFile: Record "SharePoint File" temporary; RecordID: RecordId; IsSuccess: Boolean; FileDirectory: Text; FileDirectoryFileName: Text)
    var
        E2ESharePointRec: Record "SharePoint Lists";
        SharepointSetup: Record "Sharepoint Setup";
        RecRef: RecordRef;
        FieldRef: FieldRef;
        keyRef: KeyRef;
        I: Integer;
        //Sales
        SalesHeaderRec: Record "Sales Header";
        SalesLineRec: Record "Sales Line";

    begin
        if (SharePointFile.OdataId = '') and (SharePointFile.Name = '') and ("Server Relative Url" = '') then
            exit
        else begin
            #region sharepoint file details
            if SharepointSetup.Get() then;
            E2ESharePointRec.Init();
            E2ESharePointRec.Validate("Unique Id", SharePointFile."Unique Id");
            E2ESharePointRec.Validate("Record ID", RecordID);
            E2ESharePointRec.Validate(Name, SharePointFile.Name);

            E2ESharePointRec.Insert(true);
            E2ESharePointRec.Validate(Created, SharePointFile.Created);
            E2ESharePointRec.Validate(Length, SharePointFile.Length);
            E2ESharePointRec.Validate(Exists, SharePointFile.Exists);
            E2ESharePointRec.Validate("Server Relative Url", SharePointFile."Server Relative Url");
            E2ESharePointRec.Validate(Title, SharePointFile.Title);
            E2ESharePointRec.Validate(OdataId, SharePointFile.OdataId);
            E2ESharePointRec.Validate(OdataType, SharePointFile.OdataType);
            E2ESharePointRec.Validate(OdataEditLink, SharePointFile.OdataEditLink);
            E2ESharePointRec.Validate(Id, SharePointFile.Id);
            //
            E2ESharePointRec.Validate(Link, SharepointSetup."Sharepoint Site" + SharePointFile."Server Relative Url");
            E2ESharePointRec.Validate(IsSuccess, IsSuccess);
            //
            E2ESharePointRec.Validate("Table ID", RecordID.TableNo);
            #endregion sharepoint file details

            #region table details
            Case RecordID.TableNo of
                Database::"Sales Header":
                    begin
                        if RecRef.Get(RecordID) then begin
                            RecRef.SetTable(SalesHeaderRec);
                            E2ESharePointRec.Validate("Table Name", SalesHeaderRec.TableCaption);
                            E2ESharePointRec.Validate("Document No", SalesHeaderRec."No.");
                            E2ESharePointRec.Validate("Document Type", SalesHeaderRec."Document Type");

                        end;
                    end;
                Database::"sales Line":
                    begin
                        if RecRef.Get(RecordID) then begin
                            RecRef.SetTable(SalesLineRec);
                            E2ESharePointRec.Validate("Table Name", SalesLineRec.TableCaption);
                            E2ESharePointRec.Validate("Document No", SalesLineRec."Document No.");
                            E2ESharePointRec.Validate("Document Type", SalesLineRec."Document Type");
                            E2ESharePointRec.Validate("Line No", SalesLineRec."Line No.");
                        end;
                    end;
                #endregion table details
                else begin // need to add the futher steps and i've to do testing //TODO
                    if RecRef.Get(RecordID) then begin
                        for i := 1 to RecRef.KeyIndex(1).FieldCount do begin
                            FieldRef := RecRef.KeyIndex(1).FieldIndex(i);
                            case FieldRef.Type of
                                FieldRef.Type::Option:
                                    begin
                                        E2ESharePointRec."Document Type" := FieldRef.Value;
                                    end;
                                FieldRef.Type::Code:
                                    begin
                                        if E2ESharePointRec."Document No" <> '' then
                                            E2ESharePointRec."Document No" := FieldRef.Value
                                        else
                                            Message('second key'); // add extra key 4 or 5 
                                    end;
                                FieldRef.Type::Integer:
                                    begin
                                        E2ESharePointRec."Line No" := FieldRef.Value;
                                    end;
                                FieldRef.Type::RecordId:
                                    begin

                                    end;
                            end;

                        end;
                    end;


                end;
            End;
            E2ESharePointRec.Modify(true);
        end;
    end;



    #region Purchase Archive
    //////////Start////////////Neeed to take the subcriber after the Header and Line insert. BC can Roll back the transaction but We Can't do that for sharepoint

    //Purchase Quotes ---> Purchase Quotes Archive
    //Header
    [EventSubscriber(ObjectType::Codeunit, Codeunit::ArchiveManagement, OnAfterPurchHeaderArchiveInsert, '', false, false)]
    local procedure ArchiveManagement_OnAfterPurchHeaderArchiveInsert(var PurchaseHeaderArchive: Record "Purchase Header Archive"; PurchaseHeader: Record "Purchase Header")
    var
        E2eSharepointRec: Record "SharePoint Lists";
        SharepointMgt: Codeunit "Sharepoint Management";
        DocumentInStream: InStream;
        SharepointSetupRec: Record "Sharepoint Setup";
        TempBlobCu: Codeunit "Temp Blob";
    begin
        SharepointSetupRec.Get();
        E2eSharepointRec.SetRange("Record ID", PurchaseHeader.RecordId);
        if E2eSharepointRec.FindSet() then begin
            repeat
                Clear(DocumentInStream);
                SharepointMGT.OpenFileInstream(E2eSharepointRec."Server Relative Url", DocumentInStream, TempBlobCu);
                SharepointMgt.SaveFile(SharepointSetupRec."Project Directory", E2eSharepointRec.Name, DocumentInStream, PurchaseHeaderArchive.RecordId, DocumentInStream);
            until E2eSharepointRec.Next() = 0;
        end;
    end;

    //Line
    [EventSubscriber(ObjectType::Codeunit, Codeunit::ArchiveManagement, OnAfterStorePurchLineArchive, '', false, false)]
    local procedure ArchiveManagement_OnAfterStorePurchLineArchive(var PurchHeader: Record "Purchase Header"; var PurchLine: Record "Purchase Line"; var PurchHeaderArchive: Record "Purchase Header Archive"; var PurchLineArchive: Record "Purchase Line Archive")
    var
        E2eSharepointRec: Record "SharePoint Lists";
        SharepointMgt: Codeunit "Sharepoint Management";
        DocumentInStream: InStream;
        SharepointSetupRec: Record "Sharepoint Setup";
        TempBlobCu: Codeunit "Temp Blob";
    begin
        SharepointSetupRec.Get();
        E2eSharepointRec.SetRange("Record ID", PurchLine.RecordId);
        if E2eSharepointRec.FindSet() then begin
            repeat
                Clear(DocumentInStream);
                SharepointMGT.OpenFileInstream(E2eSharepointRec."Server Relative Url", DocumentInStream, TempBlobCu);
                SharepointMgt.SaveFile(SharepointSetupRec."Project Directory", E2eSharepointRec.Name, DocumentInStream, PurchLineArchive.RecordId, DocumentInStream);
            until E2eSharepointRec.Next() = 0;
        end;
    end;

    //////////end///////////Neeed to take the subcriber after the Header and Line insert. BC can Roll back the transaction but We Can't do that for sharepoint


    // Purchase Quote to Purchase order
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Purch.-Quote to Order", OnBeforeDeletePurchQuote, '', false, false)]
    local procedure "Purch.-Quote to Order_OnBeforeDeletePurchQuote"(var QuotePurchHeader: Record "Purchase Header"; var OrderPurchHeader: Record "Purchase Header"; var IsHandled: Boolean)
    var
        E2eSharepointRec: Record "SharePoint Lists";
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
        E2eSharepointRec.SetRange("Record ID", QuotePurchHeader.RecordId);
        if E2eSharepointRec.FindSet() then begin
            repeat
                Clear(DocumentInStream);
                SharepointMGT.OpenFileInstream(E2eSharepointRec."Server Relative Url", DocumentInStream, TempBlobCu);
                SharepointMgt.SaveFile(SharepointSetupRec."Project Directory", E2eSharepointRec.Name, DocumentInStream, OrderPurchHeader.RecordId, DocumentInStream);
            until E2eSharepointRec.Next() = 0;
        end;
        //Line
        PurchQuoteLineRec.Reset();
        PurchQuoteLineRec.SetRange("Document No.", QuotePurchHeader."No.");
        PurchQuoteLineRec.SetRange("Document Type", QuotePurchHeader."Document Type");
        if PurchQuoteLineRec.FindSet() then begin
            repeat
                if PurchaseLineRec.Get(QuotePurchHeader."Document Type", QuotePurchHeader."No.", PurchQuoteLineRec."Line No.") then begin
                    E2eSharepointRec.SetRange("Record ID", PurchQuoteLineRec.RecordId);
                    if E2eSharepointRec.FindSet() then begin
                        repeat
                            Clear(DocumentInStream);
                            SharepointMGT.OpenFileInstream(E2eSharepointRec."Server Relative Url", DocumentInStream, TempBlobCu);
                            SharepointMgt.SaveFile(SharepointSetupRec."Project Directory", E2eSharepointRec.Name, DocumentInStream, PurchaseLineRec.RecordId, DocumentInStream);
                        until E2eSharepointRec.Next() = 0;
                    end;
                end;
            until PurchQuoteLineRec.Next() = 0;
        end;
    end;

    // Purchase Order ----> Purchase Order Archive 
    [EventSubscriber(ObjectType::Codeunit, Codeunit::ArchiveManagement, OnAfterStorePurchDocument, '', false, false)]
    local procedure ArchiveManagement_OnAfterStorePurchDocument(var PurchaseHeader: Record "Purchase Header"; var PurchaseHeaderArchive: Record "Purchase Header Archive")
    var
        E2eSharepointRec: Record "SharePoint Lists";
        SharepointMgt: Codeunit "Sharepoint Management";
        DocumentInStream: InStream;
        SharepointSetupRec: Record "Sharepoint Setup";
        TempBlobCu: Codeunit "Temp Blob";
        PurchaseLineArchive: Record "Purchase Line Archive";
        PurchLine: Record "Purchase Line";
    begin
        SharepointSetupRec.Get();
        //Header
        E2eSharepointRec.SetRange("Record ID", PurchaseHeader.RecordId);
        if E2eSharepointRec.FindSet() then begin
            repeat
                Clear(DocumentInStream);
                SharepointMGT.OpenFileInstream(E2eSharepointRec."Server Relative Url", DocumentInStream, TempBlobCu);
                SharepointMgt.SaveFile(SharepointSetupRec."Project Directory", E2eSharepointRec.Name, DocumentInStream, PurchaseHeaderArchive.RecordId, DocumentInStream);
            until E2eSharepointRec.Next() = 0;
        end;
        //Line
        PurchaseLineArchive.SetRange("Document Type", PurchaseHeaderArchive."Document Type");
        PurchaseLineArchive.SetRange("Document No.", PurchaseHeaderArchive."No.");
        PurchaseLineArchive.SetRange("Doc. No. Occurrence", PurchaseHeaderArchive."Doc. No. Occurrence");
        PurchaseLineArchive.SetRange("Version No.", PurchaseHeaderArchive."Version No.");
        if PurchaseLineArchive.FindSet() then begin
            // PurchaseOrderLine
            if PurchLine.Get(PurchaseHeader."Document Type", PurchaseHeader."No.", PurchaseLineArchive."Line No.") then begin
                E2eSharepointRec.SetRange("Record ID", PurchLine.RecordId);
                if E2eSharepointRec.FindSet() then begin
                    repeat
                        Clear(DocumentInStream);
                        SharepointMGT.OpenFileInstream(E2eSharepointRec."Server Relative Url", DocumentInStream, TempBlobCu);
                        SharepointMgt.SaveFile(SharepointSetupRec."Project Directory", E2eSharepointRec.Name, DocumentInStream, PurchaseLineArchive.RecordId, DocumentInStream);
                    until E2eSharepointRec.Next() = 0;
                end;
            end;
        end;
    end;
    // Purchase Order ------Purchase Order Archive
    // Dragon

    // Purchase Order ------> Posted Purchase Order Invoice
    // [EventSubscriber(ObjectType::Codeunit, Codeunit::"Purch.-Post", OnFinalizePostingOnBeforeUpdateAfterPosting, '', false, false)]
    // local procedure "Purch.-Post_OnFinalizePostingOnBeforeUpdateAfterPosting"(var PurchHeader: Record "Purchase Header"; var TempDropShptPostBuffer: Record "Drop Shpt. Post. Buffer" temporary; var EverythingInvoiced: Boolean; var IsHandled: Boolean; var TempPurchLine: Record "Purchase Line" temporary)
    // begin
    //     Message('Header no %1 and Line No %2', PurchHeader."No.", TempPurchLine."Line No.");
    // end;
    #endregion Purchase Archive


    #region Sales Archive
    // Sales Quote ----> Sales Quotes Archive
    // Header OnAfterSalesHeaderArchiveInsert
    // Line OnAfterStoreSalesLineArchive
    // Dragon
    // Sales Quote ----> Sales order
    // OnAfterInsertAllSalesOrderLines or OnBeforeDeleteSalesQuote
    // Dragon
    // Sales order ----> Sales Order Archive
    // Header OnAfterSalesHeaderArchiveInsert & Line OnAfterStoreSalesLineArchive
    // Dragon
    // Sales order ----> Posted sales invoice
    // ?
    // Dragon//     
    #endregion Sales Archive















    #region Sharepoint-Deletion
    //--------------Purchase----Start-------------------------//
    [EventSubscriber(ObjectType::Table, Database::"Purchase Header", OnAfterDeleteEvent, '', false, false)]
    local procedure OnAfterDeleteEventPurchaseHeader(var Rec: Record "Purchase Header")
    var
        E2eSharepointRec: Record "SharePoint Lists";
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
        if E2eSharepointRec.FindSet() then begin
            repeat
                SharepointMgt.DeleteFile(E2eSharepointRec."Server Relative Url");
            until E2eSharepointRec.Next() = 0;
        end;
        E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
        if E2eSharepointRec.FindSet() then
            E2eSharepointRec.DeleteAll();
    end;

    [EventSubscriber(ObjectType::Table, Database::"Purchase Line", OnAfterDeleteEvent, '', false, false)]
    local procedure OnAfterDeleteEventPurchaseLine(var Rec: Record "Purchase Line")
    var
        E2eSharepointRec: Record "SharePoint Lists";
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
        if E2eSharepointRec.FindSet() then begin
            repeat
                SharepointMgt.DeleteFile(E2eSharepointRec."Server Relative Url");
            until E2eSharepointRec.Next() = 0;
        end;
        E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
        if E2eSharepointRec.FindSet() then
            E2eSharepointRec.DeleteAll();
    end;
    //--------------Purchase----end-------------------------//

    //--------------Purchase Archive----start-------------------------//
    [EventSubscriber(ObjectType::Table, Database::"Purchase Header Archive", OnAfterDeleteEvent, '', false, false)]
    local procedure OnAfterDeleteEventPurchaseArchiveHeader(var Rec: Record "Purchase Header Archive")
    var
        E2eSharepointRec: Record "SharePoint Lists";
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
        if E2eSharepointRec.FindSet() then begin
            repeat
                SharepointMgt.DeleteFile(E2eSharepointRec."Server Relative Url");
            until E2eSharepointRec.Next() = 0;
        end;
        E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
        if E2eSharepointRec.FindSet() then
            E2eSharepointRec.DeleteAll();
    end;

    [EventSubscriber(ObjectType::Table, Database::"Purchase Line Archive", OnAfterDeleteEvent, '', false, false)]
    local procedure OnAfterDeleteEventPurchaseArchiveLine(var Rec: Record "Purchase Line Archive")
    var
        E2eSharepointRec: Record "SharePoint Lists";
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
        if E2eSharepointRec.FindSet() then begin
            repeat
                SharepointMgt.DeleteFile(E2eSharepointRec."Server Relative Url");
            until E2eSharepointRec.Next() = 0;
        end;
        E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
        if E2eSharepointRec.FindSet() then
            E2eSharepointRec.DeleteAll();
    end;
    //--------------Purchase Archive----end-------------------------//


    //--------------Sales----Start-------------------------//
    [EventSubscriber(ObjectType::Table, Database::"Sales Header", OnAfterDeleteEvent, '', false, false)]
    local procedure OnAfterDeleteEventSalesHeader(var Rec: Record "Sales Header")
    var
        E2eSharepointRec: Record "SharePoint Lists";
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
        if E2eSharepointRec.FindSet() then begin
            repeat
                SharepointMgt.DeleteFile(E2eSharepointRec."Server Relative Url");
            until E2eSharepointRec.Next() = 0;
        end;
        E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
        if E2eSharepointRec.FindSet() then
            E2eSharepointRec.DeleteAll();
    end;

    [EventSubscriber(ObjectType::Table, Database::"Sales Line", OnAfterDeleteEvent, '', false, false)]
    local procedure OnAfterDeleteEventSalesLine(var Rec: Record "Sales Line")
    var
        E2eSharepointRec: Record "SharePoint Lists";
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
        if E2eSharepointRec.FindSet() then begin
            repeat
                SharepointMgt.DeleteFile(E2eSharepointRec."Server Relative Url");
            until E2eSharepointRec.Next() = 0;
        end;
        E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
        if E2eSharepointRec.FindSet() then
            E2eSharepointRec.DeleteAll();
    end;
    //--------------Sales----End-------------------------//

    //--------------Sales Archive----Start-------------------------//
    [EventSubscriber(ObjectType::Table, Database::"Sales Header Archive", OnAfterDeleteEvent, '', false, false)]
    local procedure OnAfterDeleteEventSalesHeaderArchive(var Rec: Record "Sales Header Archive")
    var
        E2eSharepointRec: Record "SharePoint Lists";
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
        if E2eSharepointRec.FindSet() then begin
            repeat
                SharepointMgt.DeleteFile(E2eSharepointRec."Server Relative Url");
            until E2eSharepointRec.Next() = 0;
        end;
        E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
        if E2eSharepointRec.FindSet() then
            E2eSharepointRec.DeleteAll();
    end;

    [EventSubscriber(ObjectType::Table, Database::"Sales Line Archive", OnAfterDeleteEvent, '', false, false)]
    local procedure OnAfterDeleteEventSalesLineArchive(var Rec: Record "Sales Line Archive")
    var
        E2eSharepointRec: Record "SharePoint Lists";
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
        if E2eSharepointRec.FindSet() then begin
            repeat
                SharepointMgt.DeleteFile(E2eSharepointRec."Server Relative Url");
            until E2eSharepointRec.Next() = 0;
        end;
        E2eSharepointRec.SetRange("Record ID", Rec.RecordId);
        if E2eSharepointRec.FindSet() then
            E2eSharepointRec.DeleteAll();
    end;
    //--------------Sales Archive----End-------------------------//




    #endregion Sharepoint-Deletion

}