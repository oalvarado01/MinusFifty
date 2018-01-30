﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MinusFifty
{
    public struct SAttendanceEvent
    {
        public bool active;
        public string type, note;
        public List<string> members;
        public List<string> excused;

        public SAttendanceEvent(string type, string note)
        {
            this.active = true;
            this.type = type;
            this.note = note;
            this.members = new List<string>();
            this.excused = new List<string>();
        }

        public void Clear()
        {
            this.active = false;
            this.type = "";
            this.note = "";
            this.members.Clear();
            this.excused.Clear();
        }
    }
}
