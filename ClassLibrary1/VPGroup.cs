using System;
using System.Collections.Generic;


namespace CustomReport2020
{
    public class VPGroup
    {
        public string groupName { get; set; }
        public List<string> children = new List<string>();
        public List<VPGroup> subGroups = new List<VPGroup>();

        public VPGroup(string name)
        {
            this.groupName = name;
        }
    }
}
