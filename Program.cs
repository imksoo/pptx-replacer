using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

if (args.Length < 3)
{
  Console.Error.WriteLine("pptx-replacer [filename] [old-text] [new-text]");
  return;
}

string filename = args[0];
string oldText = args[1];
string newText = args[2];

if (File.Exists(filename))
{
  ReplacePptxFileText(filename, oldText, newText);
}
else
{
  Console.Error.WriteLine("File not exists: {0}", filename);
}

void ReplacePptxFileText(string filename, string oldText, string newText)
{
  using (PresentationDocument? presentationDocument = PresentationDocument.Open(filename, true))
  {
    PresentationPart? presentationPart = presentationDocument.PresentationPart;
    if (presentationPart?.Presentation?.SlideIdList != null)
    {
      foreach (SlideId slideId in presentationPart.Presentation.SlideIdList.Cast<SlideId>())
      {
        string? relationshipId = slideId?.RelationshipId;
        if (!string.IsNullOrEmpty(relationshipId))
        {
          SlidePart slidePart = (SlidePart)presentationPart.GetPartById(relationshipId);
          ReplaceSlideText(slidePart, oldText, newText);
        }
      }
    }

    presentationDocument.Save();
  }
}

void ReplaceSlideText(SlidePart slidePart, string oldText, string newText)
{
  var paragraphTextList = new List<string>();

  if (slidePart?.Slide == null) return;

  foreach (Drawing.Paragraph paragraph in slidePart.Slide.Descendants<Drawing.Paragraph>())
  {
    foreach (Drawing.Run run in paragraph.Elements<Drawing.Run>())
    {
      if (run.Text != null)
      {
        string t = run.Text.InnerText;
        if (t.Contains(oldText))
        {
          var newRunText = new Drawing.Text(t.Replace(oldText, newText));
          run.ReplaceChild(newRunText, run.Text);
        }
      }
    }
  }
}
