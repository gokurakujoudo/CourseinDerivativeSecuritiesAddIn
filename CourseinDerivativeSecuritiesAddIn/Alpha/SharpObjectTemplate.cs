using System;
using System.Collections.Generic;
using System.Linq;

namespace DerivativeSecuritiesAddIn.Alpha {
    internal class SharpObjectTemplate {
        public SharpObjectTemplate(string typeName) {
            this.TypeName = typeName.ToUpper();
            this.PropertyList = new Dictionary<string, object>();
            this.ComputedPropertyList = new Dictionary<string, Func<SharpObject, object>>();
            this.PropertyDescription = new Dictionary<string, string>();
            this.FuthurAdjust = null;
        }

        public string TypeName { get; }
        public Dictionary<string, object> PropertyList { get; }
        public Dictionary<string, string> PropertyDescription { get; }
        public Dictionary<string, Func<SharpObject, object>> ComputedPropertyList { get; }

        public Action<SharpObject> FuthurAdjust { get; }

        public bool CheckValid(SharpObject obj) => this.PropertyList.All(
            pair => obj[pair.Key] != null || pair.Value != null);

        public void Adjust(SharpObject obj) {
            foreach (var pair in this.PropertyList)
                if (obj[pair.Key] == null) obj[pair.Key] = pair.Value;
            foreach (var pair in this.ComputedPropertyList)
                obj[pair.Key] = pair.Value;
            this.FuthurAdjust?.Invoke(obj);
        }

    }
}
