pageextension 99991 "Sales Order Subform Ext" extends "Sales Order Subform" //46
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
