pageextension 50112 DocumentAttachmentDetailsExt extends "Document Attachment Details"
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
                    DocAttach: Record "Document Attachment";
                    InStr: InStream;
                    OutStr: OutStream;
                    TempBlob: Codeunit "Temp Blob";
                    DocumentReportMgt: Codeunit "Document Report Mgt.";
                    FileName: Text;
                begin
                    DocAttach.Reset();
                    CurrPage.SetSelectionFilter(DocAttach);
                    if DocAttach.FindFirst() then
                        if DocAttach."Document Reference ID".HasValue then begin
                            TempBlob.CreateOutStream(OutStr);
                            DocAttach."Document Reference ID".ExportStream(OutStr);
                            DocumentReportMgt.ConvertWordToPdf(TempBlob, 0);
                            TempBlob.CreateInStream(InStr);
                            FileName := Format(DocAttach."File Name" + '.pdf');
                            File.ViewFromStream(InStr, FileName + '.' + 'pdf', true);
                        end;
                end;
            }
        }
    }
}
