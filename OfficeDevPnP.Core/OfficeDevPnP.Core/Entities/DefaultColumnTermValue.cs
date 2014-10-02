﻿using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Entities
{
    /// <summary>
    /// Specifies a default column value for a document library
    /// </summary>
    public class DefaultColumnTermValue
    {
        /// <summary>
        /// The Path of the folder, Rootfolder of the document library is "/" 
        /// </summary>
        public string FolderRelativePath { get; set; }

        /// <summary>
        /// The internal name of the field
        /// </summary>
        public string FieldInternalName { get; set; }

        /// <summary>
        /// Taxonomy paths in the shape of "TermGroup|TermSet|Term"
        /// </summary>
        public IList<Term> Terms { get; private set; }

        public DefaultColumnTermValue()
        {
            Terms = new List<Term>();
        }
    }
}
