using EnvDTE;
using Microsoft.VisualStudio.Text;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace AlphabetizeMe
{
    [Command(PackageIds.AlphabetizeCommand)]
    internal sealed class AlphabetizeCommand : BaseCommand<AlphabetizeCommand>
    {
        #region Fields..
        private const string METHOD_PATTERN = @"(?<MethodText>(private|public|protected|internal)[\s\w<>\[\]]*?(?<MethodName>[^\s]+)\s*?\(.*?\).*?{(?>{(?<c>)|[^{}]+|}(?<-c>))*(?(c)(?!))})";
        private const string PROPERTY_PATTERN = @"(?<PropertyText>(public|private|internal|protected)(?!\(.*?\))*?[\s\w]+?(?<PropertyType>[\w<>\[\]]+)\s(?<PropertyName>[\w]+?)\s*?({(?>{(?<c>)|[^{}]+|}(?<-c>))*(?(c)(?!))}|(=[^{}]*?)*;))";

        private static Regex _methodRegex;
        private static Regex _propertyRegex;
        #endregion Fields..

        #region Structs..
        private struct Method
        {
            public string Name;
            public string Text;
            public int StartIndex;
        }

        public struct Property
        {
            public string Name;
            public string Text;
            public int StartIndex;
        }
        #endregion Structs..

        #region Constructors..
        public AlphabetizeCommand()
        {
            Initialize();
        }
        #endregion Constructors..

        #region Methods..
        private void Initialize()
        {
            _methodRegex = new Regex(METHOD_PATTERN, RegexOptions.Singleline);
            _propertyRegex = new Regex(PROPERTY_PATTERN, RegexOptions.Singleline);
        }

        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            var documentView = await VS.Documents.GetActiveDocumentViewAsync();
            var selections = documentView?.TextView.Selection.SelectedSpans.ToList() ?? new List<SnapshotSpan>();

            selections.ForEach(selection =>
            {
                var textEdit = documentView.TextBuffer.CreateEdit();

                AlphabetizeMethods(textEdit, selection);
                AlphabetizeProperties(textEdit, selection);

                textEdit.Apply();
            });
        } 

        private void AlphabetizeMethods(ITextEdit textEdit, SnapshotSpan selection)
        {
            var selectionMethods = new List<Method>();

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
            for (int i = 0; i < selectionMethods.Count; i++)
                textEdit.Replace(selectionMethods[i].StartIndex, selectionMethods[i].Text.Length, sortedSelectionMethods[i].Text);
        }

        private void AlphabetizeProperties(ITextEdit textEdit, SnapshotSpan selection)
        {
            var selectionProperties = new List<Property>();

            // Extract methods from selected text
            var matches = _propertyRegex.Matches(selection.GetText());
            foreach (Match match in matches)
            {
                var propertyName = match.Groups["PropertyName"].Value;
                var propertyText = match.Groups["PropertyText"].Value;
                var propertyIndex = match.Groups["PropertyText"].Index;

                selectionProperties.Add(new Property()
                {
                    Name = propertyName,
                    Text = propertyText,
                    StartIndex = selection.Start + propertyIndex
                });
            }

            // Sort
            var sortedSelectionProperties = selectionProperties.OrderBy(x => x.Name).ToList();

            // Replace
            for (int i = 0; i < selectionProperties.Count; i++)
                textEdit.Replace(selectionProperties[i].StartIndex, selectionProperties[i].Text.Length, sortedSelectionProperties[i].Text);
        }
        #endregion Methods..
    }
}
