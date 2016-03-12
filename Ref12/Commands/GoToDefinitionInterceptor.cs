using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using SLaks.Ref12.Services;
using System.Text.RegularExpressions;

namespace SLaks.Ref12.Commands {
	class GoToDefinitionInterceptor : CommandTargetBase<VSConstants.VSStd97CmdID> {
		readonly IEnumerable<IReferenceSourceProvider> references;
		readonly ITextDocument doc;
		readonly Dictionary<string, ISymbolResolver> resolvers = new Dictionary<string, ISymbolResolver>();

		public GoToDefinitionInterceptor(IEnumerable<IReferenceSourceProvider> references, IServiceProvider sp, IVsTextView adapter, IWpfTextView textView, ITextDocument doc) : base(adapter, textView, VSConstants.VSStd97CmdID.GotoDefn) {
			this.references = references;
			this.doc = doc;

			var dte = (DTE)sp.GetService(typeof(DTE));

			// Dev12 (VS2013) has the new simpler native API
			// Dev14, and Dev12 with Roslyn preview, will have Roslyn
			// All other versions need ParseTreeNodes
			if (new Version(dte.Version).Major > 12
			 || textView.BufferGraph.GetTextBuffers(tb => tb.ContentType.IsOfType("Roslyn Languages")).Any()) {
				RoslynAssemblyRedirector.Register();
				resolvers.Add("CSharp", CreateRoslynResolver());
				resolvers.Add("Basic", CreateRoslynResolver());
			} else {
				resolvers.Add("Basic", new VBResolver());

				if (dte.Version == "12.0")
					resolvers.Add("CSharp", new CSharp12Resolver());
				else
					resolvers.Add("CSharp", new CSharp10Resolver(dte));
			}
		}
		// This reference cannot be JITted in VS2012, so I need to wrap it in a separate method.
		static ISymbolResolver CreateRoslynResolver() { return new RoslynSymbolResolver(); }

		protected override bool Execute(VSConstants.VSStd97CmdID commandId, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut) {
            // F12 is mapped to Edit.GoToDefinition, but if press F12 to trigger it, this method will only be called once with nCmdexecopt set to 0;
            // If we use command window to trigger Edit.GoToDefinition, then this method will be called twice: first call with nCmdexecopt set 0x10003,
            // second call with nCmdexecopt set to 0. See following links for the definition of the nCmdexecopt:
            // https://msdn.microsoft.com/en-us/library/microsoft.visualstudio.ole.interop.iolecommandtarget.exec(v=vs.110).aspx
            // https://msdn.microsoft.com/en-us/library/windows/desktop/aa344051(v=vs.85).aspx
            // I decided to only call this when nCmdexecopt is 0.
            if (nCmdexecopt != 0) return false;

            ISymbolResolver resolver = null;
			SnapshotPoint? caretPoint = TextView.GetCaretPoint(s => resolvers.TryGetValue(s.ContentType.TypeName, out resolver));
			if (caretPoint == null)
				return false;

			var symbol = resolver.GetSymbolAt(doc.FilePath, caretPoint.Value);
			if (symbol == null || symbol.HasLocalSource)
				return false;

			var target = references.FirstOrDefault(r => r.AvailableAssemblies.Contains(symbol.AssemblyName));
			if (target == null)
            {
                // Open symbol without source in dnspy 2.0
                // This is a quick & dirty hack so I hardcoded dnspy path; note this will not work with dnspy 1.5 as it
                // uses different command line syntax to navigate to the symbol.
                //
                // For methods in a generic type the returned xml id will contain the actual user type name which is apparently
                // not working for DnSpy: for example, M:Foo.Bar{MyNamespace.MyType}.GetAll will not work, it should be
                // M:Foo.Bar`1.GetAll instead.
                // I have found the Roslyn source code that returns the xml id in this format:
                // http://source.roslyn.codeplex.com/#Microsoft.CodeAnalysis.CSharp/DocumentationComments/DocumentationCommentIDVisitor.PartVisitor.cs,171
                // The related imlementation is internal so it doesn't seem to be possible to override this behavior. Therefore I
                // decided to do string replace after getting the xml id.
                Debug.WriteLine("Roslyn xml id before replacing: " + symbol.XmlId);
                var xmlId = Regex.Replace(symbol.XmlId, @"\{.*?\}", m => "`" + (m.Groups[0].Value.Where(ch => ch == ',').Count() + 1));
                Debug.WriteLine("Roslyn xml id after replacing: " + xmlId);
                System.Diagnostics.Process.Start(@"c:\tools\dnspy\dnspy.exe", $"{Quoted(symbol.AssemblyPath)} --select {Quoted(xmlId)}");
                return true;
            }

			Debug.WriteLine("Ref12: Navigating to IndexID " + symbol.IndexId);

			target.Navigate(symbol);
			return true;
		}

        private static string Quoted(string text) => $"\"{text}\"";

		protected override bool IsEnabled() {
			return false;	// Always pass through to the native check
		}
	}
}
