using System;
using System.Collections.Generic;
using DerivativeSecuritiesAddIn.Utility;
using Microsoft.Office.Interop.Excel;

namespace DerivativeSecuritiesAddIn.ObjectSystem {
    internal class SharpObject {
        internal Range Init { get; }
        internal Dictionary<string, object> Properties { get; }
        internal Dictionary<string, Func<SharpObject, object>> ComputedProperties { get; }

        public SharpObject(Range init, Dictionary<string, object> value) {
            this.Init = init;
            this.Properties = new Dictionary<string, object>(value);
            this.ComputedProperties = new Dictionary<string, Func<SharpObject, object>>();
        }

        public override string ToString() => $"{this.Id()} : {this.TemplateName()} [{this.Init.ToAddress()}]";
        internal string[] ToSign() => new[] {this.Id(), this.TemplateName(), this.Init.ToAddress()};

        internal object this[string key] {
            get {
                key = key.ToUpper();
                if (this.Properties.TryGetValue(key, out var v)) return v;
                return this.ComputedProperties.TryGetValue(key, out var f) ? f(this) : null;
            }
            set {
                key = key.ToUpper();
                if (value is Func<SharpObject, object> a) {
                    this.ComputedProperties[key] = a;
                    if (this.Properties.ContainsKey(key))
                        this.Properties.Remove(key);
                }
                else {
                    this.Properties[key] = value;
                    if (this.ComputedProperties.ContainsKey(key))
                        this.ComputedProperties.Remove(key);
                }
            }
        }
    }
}
