/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 */

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Extensibility;
using Microsoft.Office.Core;
using Labradox.OneNote.Utilities;
using Application = Microsoft.Office.Interop.OneNote.Application;  // Conflicts with System.Windows.Forms

#pragma warning disable CS3003 // Type is not CLS-compliant

namespace Labradox.OneNote
{
	[ComVisible(true)]
	[Guid("354039CD-7124-49CA-A151-93AF24232EB4"), ProgId("Labradox.OneNote")]

	public class AddIn : IDTExtensibility2, IRibbonExtensibility
	{
		protected Application OneNoteApplication
		{ get; set; }

		private MainForm mainForm;

		public AddIn()
		{
		}

		/// <summary>
		/// Returns the XML in Ribbon.xml so OneNote knows how to render our ribbon
		/// </summary>
		/// <param name="RibbonID"></param>
		/// <returns></returns>
		public string GetCustomUI(string RibbonID)
		{
			return Properties.Resources.ribbon;
		}

		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>
		/// Cleanup
		/// </summary>
		/// <param name="custom"></param>
		public void OnBeginShutdown(ref Array custom)
		{
			this.mainForm?.Invoke(new Action(() =>
			{
				// close the form on the forms thread
				this.mainForm?.Close();
				this.mainForm = null;
			}));
		}

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

		public void SetOneNoteApplication(Application application)
		{
			OneNoteApplication = application;
		}

		/// <summary>
		/// Cleanup
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

		public void OnStartupComplete(ref Array custom)
		{
		}

		public async Task buttonSplitPageClicked(IRibbonControl control)
		{
			MessageBox.Show("Split page button pushed! Now we'll load up the full XML hierarchy as well as the current page XML. This may take some time.");
			ShowForm();
			return;
		}

		public async Task buttonInsertTOCClicked(IRibbonControl control)
		{
			MessageBox.Show("Table of contents (TOC) button pushed! Now we'll load up the full XML hierarchy as well as the current page XML. This may take some time.");
			ShowForm();
			return;
		}

        public async Task buttonSelectedWordClicked(IRibbonControl control)
        {
            string selection = "<not implemented>";

            MessageBox.Show("Current word: " + selection);

            return;
        }

        private void ShowForm()
		{
			this.mainForm = new MainForm(this.OneNoteApplication);
			System.Windows.Forms.Application.Run(this.mainForm);
		}

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
