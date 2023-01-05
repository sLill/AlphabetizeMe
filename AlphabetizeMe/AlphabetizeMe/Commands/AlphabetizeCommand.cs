using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Text.Tagging;
using Microsoft.VisualStudio.Utilities;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text.RegularExpressions;

namespace AlphabetizeMe
{
    [Command(PackageIds.AlphabetizeCommand)]
    internal sealed class AlphabetizeCommand : BaseCommand<AlphabetizeCommand>
    {
        #region Fields..
        private const string METHOD_PATTERN = @"(?<MethodText>(private|public|protected|internal).*?(?<MethodName>\w*?[^\b])\s*?\(.*?\).*?{(?>{(?<c>)|[^{}]+|}(?<-c>))*(?(c)(?!))})";

        private static Regex _methodRegex;
        #endregion Fields..

        #region Structs..
        private struct Method
        {
            public string Name;
            public string Text;
            public int StartIndex;
        }
        #endregion Structs..

        #region Methods..
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            _methodRegex = _methodRegex ?? new Regex(METHOD_PATTERN, RegexOptions.Singleline);

            var documentView = await VS.Documents.GetActiveDocumentViewAsync();
            var selections = documentView?.TextView.Selection.SelectedSpans.ToList() ?? new List<SnapshotSpan>();

            List<Method> selectionMethods; 
            selections.ForEach(selection =>
            {
                selectionMethods = new List<Method>();

                // Extract methods from selected text
                var matches = _methodRegex.Matches(selection.GetText());
                foreach (Match match in matches)
                {
                    var methodName = match.Groups["MethodName"].Value;
                    var methodText = match.Groups["MethodText"].Value;
                    var methodIndex = match.Groups["MethodText"].Index;

                    selectionMethods.Add(new Method() 
                    { 
                        Name = methodName, 
                        Text = methodText,
                        StartIndex = selection.Start + methodIndex
                    });
                }

                // Sort
                var sortedSelectionMethods = selectionMethods.OrderBy(x => x.Name).ToList();

                // Replace
                var textEdit = documentView.TextBuffer.CreateEdit();

                for (int i = 0; i < selectionMethods.Count; i++)
                    textEdit.Replace(selectionMethods[i].StartIndex, selectionMethods[i].Text.Length, sortedSelectionMethods[i].Text);

                textEdit.Apply();
            });
        } 
        #endregion Methods..
    }
}
