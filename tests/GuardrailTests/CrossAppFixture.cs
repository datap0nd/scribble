using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;

namespace GuardrailTests
{
    // Records calls made by the production writers. No Office installation is
    // needed in CI; this does not stand in for native Office rendering QA.
    public sealed class CrossAppFixture : DynamicObject, IEnumerable
    {
        private readonly Dictionary<string, object> _values = new Dictionary<string, object>();
        private readonly List<CrossAppFixture> _items = new List<CrossAppFixture>();
        private readonly string _path;
        private readonly List<string> _events;
        public CrossAppFixture(string path, List<string> events) { _path = path; _events = events; }
        private CrossAppFixture Child(string name)
        {
            object value;
            if (!_values.TryGetValue(name, out value)) _values[name] = value = new CrossAppFixture(_path + "." + name, _events);
            return (CrossAppFixture)value;
        }
        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            if (_values.TryGetValue(binder.Name, out result)) return true;
            if (binder.Name == "ActiveWorkbook" || binder.Name == "ActiveDocument" || binder.Name == "ActivePresentation")
                throw new InvalidOperationException("Cross-app writer touched an existing document: " + binder.Name);
            switch (binder.Name)
            {
                case "Count": result = _items.Count; return true;
                case "Id": case "SlideIndex": case "End": result = 1; return true;
                case "BoundHeight": case "BoundWidth": result = 1d; return true;
                case "Text": case "Name": case "Path": case "FullName": result = ""; return true;
                case "HasTextFrame": result = -1; return true;
                case "Worksheets":
                    var sheets = Child(binder.Name);
                    if (sheets._items.Count == 0) sheets._items.Add(new CrossAppFixture(sheets._path + "[1]", _events));
                    result = sheets; return true;
                default: result = Child(binder.Name); return true;
            }
        }
        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            _values[binder.Name] = value;
            var grid = value as object[,];
            _events.Add(_path + "." + binder.Name + "=" + (grid == null ? Convert.ToString(value) : string.Join("|", grid.Cast<object>())));
            return true;
        }
        public override bool TryGetIndex(GetIndexBinder binder, object[] indexes, out object result)
        {
            var key = string.Join(",", indexes.Select(Convert.ToString));
            if (_values.TryGetValue(key, out result)) return true;
            if (indexes.Length == 1 && indexes[0] is int && (int)indexes[0] > 0 && (int)indexes[0] <= _items.Count)
                result = _items[(int)indexes[0] - 1];
            else result = Child("[" + key + "]");
            return true;
        }
        public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
        {
            _events.Add(_path + "." + binder.Name + "(" + string.Join(",", args.Select(Convert.ToString)) + ")");
            if (new[] { "Save", "SaveAs", "Send", "Close", "Quit", "Delete" }.Contains(binder.Name))
                throw new InvalidOperationException("Forbidden operation: " + binder.Name);
            if (binder.Name == "Export")
            {
                using (var bitmap = new Bitmap(160, 90)) bitmap.Save((string)args[0], ImageFormat.Png);
                result = null; return true;
            }
            if (binder.Name == "Add" && _path.EndsWith(".Tags"))
            { _values[(string)args[0]] = args[1]; result = null; return true; }
            if (binder.Name == "Add" || binder.Name.StartsWith("AddText") || binder.Name == "AddShape")
            {
                var child = new CrossAppFixture(_path + "[" + (_items.Count + 1) + "]", _events);
                child._values["Id"] = _items.Count + 1;
                child._values["SlideIndex"] = _items.Count + 1;
                _items.Add(child); result = child; return true;
            }
            result = Child(binder.Name); return true;
        }
        public IEnumerator GetEnumerator() { return _items.GetEnumerator(); }
    }
}
