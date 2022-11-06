using Microsoft.Office.Interop.Word;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
    

namespace Invoice_Initializer
{
    public class FindAndReplace
    {
        public void FindReplace(Application wordApp, object FindText, object ReplaceWith)
        {

            object MatchCase = true;
            object MatchWholeWord = true;
            object MatchWildcards = false;
            object MatchSoundsLike = false;
            object MatchAllWordForms = false;
            object Forward = true;
            object Wrap = 1;
            object Format = false;
            object Replace = WdReplace.wdReplaceAll;
            object MatchKashida = WdFindWrap.wdFindContinue;
            object MatchDiacritics = false;
            object MatchControl = false;
            object MatchAlefHamza = false;

            wordApp.Selection.Find.Execute(ref FindText,
                ref MatchCase,
                ref MatchWholeWord,
                ref MatchWildcards,
                ref MatchSoundsLike,
                ref MatchAllWordForms,
                ref Forward,
                ref Wrap,
                ref Format,
                ref ReplaceWith,
                ref Replace,
                ref MatchKashida,
                ref MatchDiacritics,
                ref MatchControl,
                ref MatchAlefHamza);
        }

        public void ReplaceHeader(Application wordApp, object FindText, object ReplaceWith, Word.Document myWordDoc)
        {
            object wdFindContinue = (object)Word.WdFindWrap.wdFindContinue;
            object wdReplaceAll = (object)Word.WdReplace.wdReplaceAll;
            object missing = Missing.Value;

            wordApp.Selection.Find.ClearFormatting();

            wordApp.Selection.Find.Replacement.ClearFormatting();


            var junk = myWordDoc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.StoryType;

            foreach (Word.Range range in myWordDoc.StoryRanges)
            {
                Word.Range range2 = range;

                while (range2 != null)
                {
                    range2.Find.Execute(ref FindText, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref wdFindContinue,
                    ref missing, ref ReplaceWith, ref wdReplaceAll, ref missing, ref missing, ref missing, ref missing);

                    foreach (Word.Shape shape in range2.ShapeRange)
                    {
                        if (shape.TextFrame.HasText != 0)
                        {
                            shape.TextFrame.TextRange.Find.Execute(ref FindText, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref wdFindContinue,
                            ref missing, ref ReplaceWith, ref wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                        }
                    }

                    if(range2.StoryType == WdStoryType.wdPrimaryHeaderStory)
                    {
                        if (range.ShapeRange.Count > 0)
                        {
                            foreach (Word.Shape shape in range2.ShapeRange)
                            {
                                if (shape.TextFrame.HasText != 0)
                                {
                                    shape.TextFrame.TextRange.Find.Execute(ref FindText, ref missing, ref missing, ref missing, ref missing,
                                                                            ref missing, ref missing, ref wdFindContinue,
                                                                            ref missing, ref ReplaceWith, ref wdReplaceAll,
                                                                            ref missing, ref missing, ref missing, ref missing);
                                }
                            }
                        }
                    }
                    range2 = range2.NextStoryRange;
                }
            }
        }
    }
}
