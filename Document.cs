namespace Manage_Document
{
    using System;
    using System.Collections.Generic;
    
    public partial class Document
    {
        public int ID { get; set; }
        public string Ten { get; set; }
        public string LinkImage { get; set; }
        public string Link { get; set; }
        public Nullable<int> IsRead { get; set; }
    }
}
