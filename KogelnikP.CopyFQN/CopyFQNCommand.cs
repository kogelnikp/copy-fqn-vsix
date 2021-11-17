using Microsoft;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.ComponentModel.Design;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace KogelnikP.CopyFQN
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class CopyFQNCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("039ba84e-af40-40a2-a032-d650fb0e4443");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="CopyFQNCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private CopyFQNCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new OleMenuCommand(this.Execute, menuCommandID);
            menuItem.BeforeQueryStatus += MenuItem_BeforeQueryStatus;
            commandService.AddCommand(menuItem);
        }



        /// <summary>
        /// Handler for the BeforeQueryStatus event.
        /// Checks whether the command button should be visible.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            if (sender is OleMenuCommand command)
            {
                command.Visible = false;
                ThreadHelper.JoinableTaskFactory.Run(async delegate
                {
                    (var node, var document) = await GetCurrentTokenAndDocumentAsync();
                    command.Visible = IsSyntaxNodeSupported(node);
                });
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static CopyFQNCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in CopyFQNCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new CopyFQNCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            ThreadHelper.JoinableTaskFactory.Run(async delegate {
                (var node, var document) = await GetCurrentTokenAndDocumentAsync();

                if (IsSyntaxNodeSupported(node))
                {
                    SemanticModel semanticModel = await document.GetSemanticModelAsync();
                    var symbol = semanticModel.GetDeclaredSymbol(node);
                    var containingNamespace = symbol.ContainingNamespace;

                    System.Windows.Forms.Clipboard.SetText(symbol.ToDisplayString(SymbolDisplayFormat.CSharpErrorMessageFormat));
                }

            });
        }


        /// <summary>
        /// Checks if the syntax node type is supported
        /// </summary>
        /// <param name="node"></param>
        /// <returns></returns>
        private bool IsSyntaxNodeSupported(SyntaxNode node)
        {
            return node != null 
                && ((node is MemberDeclarationSyntax));
        }


        /// <summary>
        /// Retrieves 
        /// </summary>
        /// <returns></returns>
        private async Task<(SyntaxNode, Document)> GetCurrentTokenAndDocumentAsync()
        {
            IWpfTextView textView = GetTextView();
            SnapshotPoint caretPosition = textView.Caret.Position.BufferPosition;
            Document document = caretPosition.Snapshot.GetOpenDocumentInCurrentContextWithChanges();
            var syntaxRoot = await document.GetSyntaxRootAsync();
            return (syntaxRoot.FindToken(caretPosition).Parent, document);
        }


        /// <summary>
        /// Gets the current editor view
        /// </summary>
        /// <returns></returns>
        private IWpfTextView GetTextView()
        {
            IVsTextManager textManager =
                (IVsTextManager)ServiceProvider.GetService(
                    typeof(SVsTextManager));
            Assumes.Present(textManager);
            textManager.GetActiveView(1, null, out IVsTextView textView);
            return (GetEditorAdaptersFactoryService()).GetWpfTextView(textView);
        }


        /// <summary>
        /// Gets the EditorAdaptersFactoryService
        /// </summary>
        /// <returns></returns>
        private IVsEditorAdaptersFactoryService GetEditorAdaptersFactoryService()
        {
            IComponentModel componentModel =
                (IComponentModel)ServiceProvider.GetService(
                    typeof(SComponentModel));
            Assumes.Present(componentModel);
            return componentModel.GetService<IVsEditorAdaptersFactoryService>();
        }
    }
}
