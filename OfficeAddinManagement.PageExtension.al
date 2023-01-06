pageextension 50100 "Office Add-in Management Ext" extends "Office Add-in Management"
{

    actions
    {
        addlast(processing)
        {
            action("Custom Outlook Actions")
            {
                Caption = 'Custom Outlook Actions';
                RunObject = page "Custom Outlook Actions";
                Image = AddAction;
            }
            action("Document Viewer")
            {
                Caption = 'Document Viewer';
                RunObject = page "Outl. Document Viewer";
                Image = AddAction;
            }
        }
    }
}