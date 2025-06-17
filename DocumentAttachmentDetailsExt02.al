pageextension 50110 DocumentAttachmentDetailsExt extends "Document Attachment Details"
{
    actions
    {
        addafter(Preview_Promoted)
        {
            actionref(PreviewWord_Promoted; PreviewWord)
            {
            }
        }
        addfirst(processing)
        {
            action(PreviewWord)
            {
                Caption = 'Preview Word';
                ApplicationArea = All;
                Image = View;
                trigger OnAction()
                var
                    PreviewFiles: Page "Preview Files";
                    DocAttach: Record "Document Attachment";
                    HTMLText: Text;
                    InStr: InStream;
                    OutStr: OutStream;
                    TempBlob: Codeunit "Temp Blob";
                    DocumentReportMgt: Codeunit "Document Report Mgt.";
                begin
                    DocAttach.Reset();
                    CurrPage.SetSelectionFilter(DocAttach);
                    if DocAttach.FindFirst() then
                        if DocAttach."Document Reference ID".HasValue then begin
                            TempBlob.CreateOutStream(OutStr);
                            DocAttach."Document Reference ID".ExportStream(OutStr);
                            DocumentReportMgt.ConvertWordToHtml(TempBlob);
                            TempBlob.CreateInStream(InStr);
                            InStr.ReadText(HTMLText);
                            PreviewFiles.SetURL(HTMLText);
                            PreviewFiles.Run();
                        end;
                end;
            }
        }
    }
}
page 50111 "Preview Files"
{
    Caption = 'Preview';
    Editable = false;
    PageType = Card; //Use new UserControlHost PageType from BC26
    layout
    {
        area(content)
        {
            usercontrol(WebPageViewer; WebPageViewer)
            {
                ApplicationArea = All;
                trigger ControlAddInReady(callbackUrl: Text)
                begin
                    CurrPage.WebPageViewer.SetContent(HTMLText);
                end;

                trigger Callback(data: Text)
                begin
                    CurrPage.Close();
                end;
            }
        }
    }
    var
        HTMLText: Text;

    procedure SetURL(NavigateToURL: Text)
    begin
        HTMLText := NavigateToURL;
    end;
}
