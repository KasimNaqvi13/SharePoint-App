// pageextension 50159 "Sales Order Subform Ext" extends "Sales Order Subform" //46
// {
//     layout
//     {
//         modify(ShortcutDimCode3)
//         {
//             Visible = false;
//         }
//         modify(ShortcutDimCode4)
//         {
//             Visible = false;
//         }
//         modify(ShortcutDimCode5)
//         {
//             Visible = false;
//         }
//         modify(ShortcutDimCode6)
//         {
//             Visible = false;
//         }
//         modify(ShortcutDimCode7)
//         {
//             Visible = false;
//         }
//         modify(ShortcutDimCode8)
//         {
//             Visible = false;
//         }
//     }

//     actions
//     {
//         addafter("Item Charge &Assignment")
//         {

//             action("Share-Point Attachment")
//             {
//                 Caption = 'Share Point Attachment';
//                 ApplicationArea = all;
//                 Image = Attachments;
//                 // RunObject = page "E2E-SharePoint Subform";
//                 // RunPageLink = "Document No" = field("Document No."), "Document Type" = field("Document Type"), "Line No" = field("Line No.");
//                 trigger OnAction()
//                 var
//                     E2ESharepointSubform: Page "SharePoint Subform";
//                     E2ESharepointRec: Record "SharePoint Lists";
//                 begin
//                     E2ESharepointRec.SetRange("Document No", Rec."Document No.");
//                     E2ESharepointRec.SetRange("Document Type", rec."Document Type");
//                     E2ESharepointRec.SetRange("Line No", Rec."Line No.");
//                     E2ESharepointRec.SetRange("Record ID", Rec.RecordId);

//                     E2ESharepointSubform.SetRecord(Rec.RecordId, true);
//                     E2ESharepointSubform.SetTableView(E2ESharepointRec);
//                     E2ESharepointSubform.RunModal()

//                 end;
//             }
//         }

//     }
// }
