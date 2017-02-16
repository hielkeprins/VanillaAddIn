/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 */

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Drawing.Imaging;
using System.IO;
using System.Xml;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

using Extensibility;
using Microsoft.Office.Core;

using Labradox.OneNote.Utilities;

// Conflicts with System.Windows.Forms
using Application = Microsoft.Office.Interop.OneNote.Application;
using Microsoft.Office.Interop.OneNote;
#pragma warning disable CS3003 // Type is not CLS-compliant

namespace Labradox.OneNote.Toolbar
{
	[ComVisible(true)]
	[Guid("354039CD-7124-49CA-A151-93AF24232EB4"), ProgId("Labradox.OneNote")]

	public class AddIn : IDTExtensibility2, IRibbonExtensibility
	{
        protected Application OneNoteApplication { get; set; }

        public AddIn() {}

        public void SetOneNoteApplication(Application application)
        {
            OneNoteApplication = application;
        }

        #region IDTExtensibility implementation (event listeners)

        public void OnAddInsUpdate(ref Array custom) {}

		public void OnBeginShutdown(ref Array custom) {}

		/// <summary>
		/// Called upon startup.
		/// Keeps a reference to the current OneNote application object.
		/// </summary>
		/// <param name="application"></param>
		/// <param name="connectMode"></param>
		/// <param name="addInInst"></param>
		/// <param name="custom"></param>
		public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
		{
			SetOneNoteApplication((Application)Application);
		}

		/// <summary>
		/// Cleanup. Disposes references to the OneNote application and triggers garbage collection.
		/// </summary>
		/// <param name="RemoveMode"></param>
		/// <param name="custom"></param>
		[SuppressMessage("Microsoft.Reliability", "CA2001:AvoidCallingProblematicMethods", MessageId = "System.GC.Collect")]
		public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
		{
			OneNoteApplication = null;
			GC.Collect();
			GC.WaitForPendingFinalizers();
		}

		public void OnStartupComplete(ref Array custom) {}

        #endregion

        #region IRibbonExtensibility implementation (GUI factory)

        /// <summary>
        /// Returns the XML in Ribbon.xml so OneNote knows how to render our ribbon
        /// </summary>
        /// <param name="RibbonID"></param>
        /// <returns></returns>
        public string GetCustomUI(string RibbonID)
        {
            return Properties.Resources.ribbon;
        }

        #endregion

        #region Ribbon callback functions

        public async Task buttonSplitPageClicked(IRibbonControl control)
        {
            return;
        }

        public async Task buttonInsertTOCClicked(IRibbonControl control)
        {
            return;
        }

        public async Task buttonPublishClicked(IRibbonControl control)
        {
            MessageBox.Show("Creating Notebook parser...");

            string xmlHierarchy;

            // get currently active notebook from OneNote
            var nbId = OneNoteApplication.Windows.CurrentWindow.CurrentNotebookId;
            OneNoteApplication.GetHierarchy(nbId, HierarchyScope.hsPages, out xmlHierarchy);

            NotebookParser nb = new NotebookParser(xmlHierarchy);

            MessageBox.Show(string.Format("Parsed {0} pages (organized in {2} sections) of notebook {1}.", nb.pages.Count, nb.displayname, nb.sections.Count));
            
            return;
        }

        public async Task buttonSelectedWordClicked(IRibbonControl control)
        {
            string selection = "<not implemented>";
            string xml;

            // get current location in notebook collection from OneNote
            var pageId = OneNoteApplication.Windows.CurrentWindow.CurrentPageId;
            var sectionId = OneNoteApplication.Windows.CurrentWindow.CurrentSectionId;
            var groupId = OneNoteApplication.Windows.CurrentWindow.CurrentSectionGroupId;
            var notebookId = OneNoteApplication.Windows.CurrentWindow.CurrentNotebookId;

            // get raw XML content of current page
            OneNoteApplication.GetPageContent(pageId, out xml);

            // parse XML to find members of atomic page element (eg. headed paragraphs)

            // loop to these elements
            // create a new page for each of these elements
            var childPageId = pageId;
            OneNoteApplication.CreateNewPage(sectionId, out childPageId, NewPageStyle.npsBlankPageWithTitle);
            // move title
            // set/move timestamp
            // move contents
            // end
            OneNoteApplication.GetHierarchy(null, HierarchyScope.hsPages, out xml);

            MessageBox.Show("Current word: " + selection);

            return;
        }

        #endregion

        /// <summary>
        /// Specified in Ribbon.xml, this method returns the image to display on the ribbon button
        /// </summary>
        /// <param name="imageName"></param>
        /// <returns></returns>
        public IStream GetImage(string imageName)
		{
			MemoryStream imageStream = new MemoryStream();
			switch (imageName)
			{
				case "ToolbarButtonSplit":
					Properties.Resources.ToolbarButtonSplit.Save(imageStream, ImageFormat.Png);
					break;
				case "ToolbarButtonTOC":
					Properties.Resources.ToolbarButtonTOC.Save(imageStream, ImageFormat.Png);
					break;

				default:
					Properties.Resources.LabradoxLogo.Save(imageStream, ImageFormat.Png);
					break;
			}
			return new CCOMStreamWrapper(imageStream);
		}
	}
}
