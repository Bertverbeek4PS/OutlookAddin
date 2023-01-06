page 50110 "Custom Outlook Actions"
{
    Caption = 'Custom Outlook Actions';
    PageType = List;
    SourceTable = "Custom Outlook Action";

    layout
    {
        area(content)
        {
            repeater(Control1100527601)
            {
                ShowCaption = false;
                field(Name; rec.Name)
                {
                    ToolTip = 'Provide the name of the action in the Outlook add-in without any special characters';
                }
                field(Label; rec.Label)
                {
                    ToolTip = 'Provide the label of the action in the Outlook add-in';
                }
                field(Tip; rec.Tip)
                {
                    ToolTip = 'Provide the yip of the action in the Outlook add-in';
                }
                field("Super Tip Title"; rec."Super Tip Title")
                {
                    ToolTip = 'Provide the tip title of the action in the Outlook add-in';
                }
                field("Super Tip Desc"; rec."Super Tip Desc")
                {
                    ToolTip = 'Provide the super tip description of the action in the Outlook add-in';
                }
                field("Page ID"; rec."Page ID")
                {
                    TableRelation = AllObjWithCaption."Object ID" where("Object Type" = Const(Page));
                    ToolTip = 'Provide the page ID which will open in the Outlook add-in';
                }
                field("Page Caption"; rec."Page Caption")
                {
                    ToolTip = 'Page Caption of the Page ID';
                }
                field("Table ID"; rec."Table ID")
                {
                    TableRelation = AllObjWithCaption."Object ID" where("Object Type" = Const(Table), "Object Subtype" = Const('Normal'));
                    ToolTip = 'Provide the table ID which will create the record from the Outlook add-in';
                }
                field("Table Caption"; rec."Table Caption")
                {
                    ToolTip = 'Table Caption of the Table ID';
                }
                field("Contact Type"; rec."Contact Type")
                {
                    ToolTip = 'The contact type of the chosen table.';
                }
                field("Field No. Contact Type"; rec."Field No. Contact Type")
                {
                    ToolTip = 'Select the contact field of the table to insert the customer, vendor or contact.';

                    trigger OnLookup(var Text: Text): Boolean
                    var
                        Fld: Record "Field";
                        FieldSelection: Codeunit "Field Selection";
                    begin
                        Fld.SetRange(TableNo, Rec."Table ID");
                        if FieldSelection.Open(Fld) then begin
                            Rec.Validate("Field No. Contact Type", Fld."No.");
                        end;
                    end;

                }
                field("Field No. Name"; rec."Field No. Name")
                {
                    ToolTip = 'Select the name field of the table to insert the customer, vendor or contact.';

                    trigger OnLookup(var Text: Text): Boolean
                    var
                        Fld: Record "Field";
                        FieldSelection: Codeunit "Field Selection";
                    begin
                        Fld.SetRange(TableNo, Rec."Table ID");
                        if FieldSelection.Open(Fld) then begin
                            Rec.Validate("Field No. Name", Fld."No.");
                        end;
                    end;
                }
                field("Field No. E-Mail"; rec."Field No. E-mail")
                {
                    ToolTip = 'Select the e-mail field of the table to insert the customer, vendor or contact.';

                    trigger OnLookup(var Text: Text): Boolean
                    var
                        Fld: Record "Field";
                        FieldSelection: Codeunit "Field Selection";
                    begin
                        Fld.SetRange(TableNo, Rec."Table ID");
                        if FieldSelection.Open(Fld) then begin
                            Rec.Validate("Field No. E-Mail", Fld."No.");
                        end;
                    end;
                }
            }
        }
    }
}

