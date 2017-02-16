using System;

using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

// for Microsoft Dev Example
using System.Linq;
using System.Xml.Linq;

using System.Collections.Generic;

using Microsoft.Office.Interop.OneNote;

namespace Labradox.OneNote.Utilities
{
    public class NotebookParser
    {
        #region Wrapper classes for xml nodes

        public class NotebookNode
        {
            public Dictionary<string, string> header { get; protected set; }

            public string id { get; protected set; }

            public NotebookNode(XmlReader node)
            {
                header["ID"] = node.GetAttribute("ID");
                header["name"] = node.GetAttribute("name");
                header["slug"] = UrlSlugger.ToUrlSlug(header["name"]);
            }
        }

        public class NotebookSection : NotebookNode
        {
            public NotebookSection(XmlReader node) : base(node)
            {
            }
        }

        public class NotebookPage : NotebookNode
        {
            public string xml { get; protected set; }

            public NotebookPage(XmlReader node) : base(node)
            {

            }
        }

        #endregion

        #region Publicly accesable data members

        /* exposed yaml data for file generation and further processing */

        // keeps track of section node attributes 
        public List<NotebookSection> sections { get; protected set; }
        // keeps track of page node attributes 
        public List<NotebookPage> pages { get; protected set; }

        // exposed raw xml of entire outline in a Notebook node
        public string xml { get; protected set; }
        // exposed display name parsed from Notebook node attribute
        public string displayname { get; protected set; }

        #endregion

        // used by private parser functions to traverse <one:Notebook> xml hierarchy
        private XmlReader hierarchy;

        private Application onenote;

        #region Constructors

        public NotebookParser() : this(string.Empty) {}

        public NotebookParser(string xml)
        {
            // reference to OneNote application
            try
            {
                onenote = new Application();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return;
            }

            // storage for xml node attributes
            pages = new List<NotebookPage>();
            sections = new List<NotebookSection>();

            if (string.IsNullOrEmpty(xml))
            {
                // by default, get xml from currently active notebook in OneNote
                var notebookId = onenote.Windows.CurrentWindow.CurrentNotebookId;
                onenote.GetHierarchy(notebookId, HierarchyScope.hsPages, out xml);
            }
            this.xml = xml;

            hierarchy = XmlReader.Create(this.xml);
        }

        #endregion

        #region Public member functions

        public void Parse()
        {
            ParseNotebookNode();
        }

        private static void MicrosoftDevExample()
        {
            // skipping error checking, just demonstrating using these APIs
            var app = new Application();

            // get the hierarchy
            string xmlHierarchy;
            app.GetHierarchy(null, HierarchyScope.hsPages, out xmlHierarchy);

            Console.WriteLine("Hierarchy:\n" + xmlHierarchy);

            if (app.Windows.CurrentWindow != null)
            {
                // now find the current page, print out its ID
                var xdoc = XDocument.Parse(xmlHierarchy);
                var ns = xdoc.Root.Name.Namespace;

                var pageId = app.Windows.CurrentWindow.CurrentPageId;
                Console.WriteLine("Current Page ID: " + pageId);

                // get the page content, print it out
                string xmlPage;
                app.GetPageContent(pageId, out xmlPage);
                Console.WriteLine("Page XML:\n" + xmlPage);

                // sample - this is how to update content - normally you would modify the xml.
                app.UpdatePageContent(xmlPage);

                // bonus - if there are any images, get the binary content of the first one
                var xPage = XDocument.Parse(xmlPage);
                var xImage = xPage.Descendants(ns + "Image").FirstOrDefault();
                if (xImage != null)
                {
                    var xImageCallbackID = xImage.Elements(ns + "CallbackID").First();
                    var imageId = xImageCallbackID.Attribute("callbackID").Value;
                    string base64Out;
                    app.GetBinaryPageContent(pageId, imageId, out base64Out);

                    Console.WriteLine("Image found, base64 data is:\n" + base64Out);
                }
            }
        }

        #endregion

        #region Private node parsing functions

        private void ParseNotebookNode()
        {
            // read data from Notebook rootnode attributes
            if (hierarchy.ReadToDescendant("one:Notebook"))
            {
                displayname = hierarchy.GetAttribute("nickname");

                ParseSectionNodes();
            }
        }

        private void ParseSectionNodes()
        {
            // skip to first section in Notebook subtree
            if (hierarchy.ReadToDescendant("one:Section"))
            {
                // traverse over all sections in the notebook
                do
                {
                    // parse pages in current section
                    ParsePageNodes(hierarchy.ReadSubtree());

                    sections.Add(new NotebookSection(hierarchy));
                }
                while (hierarchy.ReadToNextSibling("one:Section"));
            }
        }

        private void ParsePageNodes(XmlReader subtree)
        {
            // skip to first page in section
            if (subtree.ReadToDescendant("one:Page"))
            { 
                // traverse over all pages (on all levels),
                do
                {
                    pages.Add(new NotebookPage(subtree));
                }
                while (subtree.ReadToNextSibling("one:Page"));
            }
        }

        #endregion
    }

    public class NotebookGenerator : NotebookParser
    {
        // specifies the folder to for the generator output
        public string jekyllroot { get; set; } =
            @"C:\Users\bsms4079\Documents\GitHub\labradox-onenote\Setup\Debug\";

        public string collectionname { get; set; } =
            "notes";

        protected string path;

        [Flags]
        public enum PageType
        {
            yaml,
            xml,
            txt
        }

        public NotebookGenerator(string xml) : base(xml)
        {
            path = string.Format(@"{0}\_{1}\{2}", jekyllroot, collectionname, displayname);
        }

        #region Jekyll generators

        /// <summary>
        /// Stores the entire xml hierarchy in a file called after the Notebook's displayname.
        /// </summary>
        public void WriteXML()
        {
            File.WriteAllText(string.Format(@"{0}\{1}", path, "notebook.xml"), this.xml);
        }

        /// <summary>
        /// Creates a subfolder for each section in the nootbook in the output directory.
        /// </summary>
        public void CreateDirectoryStructure()
        {
            foreach (NotebookSection s in sections)
            {
                string dirname = string.Format(@"{0}\{1}", path, s.header["slug"]);

                Directory.CreateDirectory(dirname);
            }
        }

        /// <summary>
        /// Generates files with the contents of each page in the Notebook's sections in the appropriate subfolder.
        /// </summary>
        public void WriteYAML()
        {
            foreach (NotebookPage p in pages)
            {
                string pSectionId = p.header["ID"].Substring(2, 16);
                NotebookSection pSection = sections.Find(s => s.header["ID"].Equals(pSectionId));


                string filename = string.Format(@"{0}\{1}\{2}", path, pSection.header["slug"], p.header["slug"]);
                    
                using (StreamWriter yamlfile = File.CreateText(string.Format("%s.yaml", filename)))
                {
                    yamlfile.WriteLine("---");
                    foreach (KeyValuePair<string, string> yaml in p.header)
                    {
                        yamlfile.Write(string.Format(@"{0}: {1}", yaml.Key, yaml.Value));
                    }
                    yamlfile.WriteLine("---");
                }
            }
        }

        public void GetSection(string pageId)
        {

        }

        /*
        public void PublishPages()
        {
            foreach (NotebookSection s in sections)
            {
                foreach (NotebookPage p in pages)
                {
                    string pageXml;

                    app.GetPageContent(p.header["ID"], out pageXml);

                    PublishPage(pageXml);
                }
            }
        }

        /// <summary>
        /// Generates files with the contents of each page in the Notebook's sections in the appropriate subfolder.
        /// </summary>
        public void PublishPage(string pageXml)
        {         
            string filename = string.Format(@"{0}\{1}\{2}", path, s.header["slug"], p.header["slug"]);

            if (!string.IsNullOrEmpty(pageXml))
            {
                // Raw XML
                if (true) // (type.HasFlag(PageType.xml))
                {
                    File.WriteAllText(filename + ".xml", pageXml);
                }

                        
                // Barebone Markdown subset
                if (true) // (type.HasFlag(PageType.txt))
                {
                    File.WriteAllText(filename + ".txt", pageXml);
                }
            }
        }
        */

        #endregion
    }
}