pageextension 99991 "Purchase Order Subform Ext" extends "Purchase Order Subform" //54
{
    actions
    {
        addlast(Page)
        {
            action("Share-Point Attachment")
            {
                Caption = 'Share Point Attachment';
                ApplicationArea = all;
                Image = Attachments;
                trigger OnAction()
                var
                    SharePointListLogSubPage: Page "SharePoint List Log Subform";
                    SharePointListLog: Record "SharePoint List Log";
                begin
                    SharePointListLog.SetRange("Table ID", Database::"Sales Line");
                    SharePointListLog.SetRange("Table Record ID", Rec.RecordId);
                    SharePointListLogSubPage.SetRecord(Rec.RecordId, true);
                    SharePointListLogSubPage.SetTableView(SharePointListLog);
                    SharePointListLogSubPage.RunModal();
                end;
            }
        }
    }
}
