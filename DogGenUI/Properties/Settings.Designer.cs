﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace DocGeneratorCore.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "14.0.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.WebServiceUrl)]
        [global::System.Configuration.DefaultSettingValueAttribute("https://teams.dimensiondata.com/_vti_bin/copy.asmx")]
        public string SDDPwebReferencePROD {
            get {
                return ((string)(this["SDDPwebReferencePROD"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("DocGenerator\\Database\\Production")]
        public string DatabaseLocationPROD {
            get {
                return ((string)(this["DatabaseLocationPROD"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("DocGenerator\\Database\\QualityAssurance")]
        public string DatabaseLocationQA {
            get {
                return ((string)(this["DatabaseLocationQA"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("\\Database\\License")]
        public string DatabaseLocationLicense {
            get {
                return ((string)(this["DatabaseLocationLicense"]));
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string CurrentDatabaseLocation {
            get {
                return ((string)(this["CurrentDatabaseLocation"]));
            }
            set {
                this["CurrentDatabaseLocation"] = value;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("DocGenerator\\Database\\Development")]
        public string DatabaseLocationDEV {
            get {
                return ((string)(this["DatabaseLocationDEV"]));
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string CurrentURLSharePoint {
            get {
                return ((string)(this["CurrentURLSharePoint"]));
            }
            set {
                this["CurrentURLSharePoint"] = value;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("https://teams.dimensiondata.com")]
        public string URLSharePointPROD {
            get {
                return ((string)(this["URLSharePointPROD"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("https://teams.uat.dimensiondata.com")]
        public string URLSharePointDEV {
            get {
                return ((string)(this["URLSharePointDEV"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("https://teams.dimensiondata.com")]
        public string URLSharePointQA {
            get {
                return ((string)(this["URLSharePointQA"]));
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string CurrentDatabaseHost {
            get {
                return ((string)(this["CurrentDatabaseHost"]));
            }
            set {
                this["CurrentDatabaseHost"] = value;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.WebServiceUrl)]
        [global::System.Configuration.DefaultSettingValueAttribute("https://teams.dimensiondata.com/_vti_bin/copy.asmx")]
        public string SDDPwebReferenceQA {
            get {
                return ((string)(this["SDDPwebReferenceQA"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.WebServiceUrl)]
        [global::System.Configuration.DefaultSettingValueAttribute("https://teams.uat.dimensiondata.com/_vti_bin/copy.asmx")]
        public string SDDPwebReferenceDEV {
            get {
                return ((string)(this["SDDPwebReferenceDEV"]));
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string CurrentSDDPwebReference {
            get {
                return ((string)(this["CurrentSDDPwebReference"]));
            }
            set {
                this["CurrentSDDPwebReference"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string CurrentPlatform {
            get {
                return ((string)(this["CurrentPlatform"]));
            }
            set {
                this["CurrentPlatform"] = value;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("/sites/servicecatalogue")]
        public string URLSharePointSitePortionPROD {
            get {
                return ((string)(this["URLSharePointSitePortionPROD"]));
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string CurrentURLSharePointSitePortion {
            get {
                return ((string)(this["CurrentURLSharePointSitePortion"]));
            }
            set {
                this["CurrentURLSharePointSitePortion"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        public global::System.DateTime CurrentDatabaseLastRefreshedOn {
            get {
                return ((global::System.DateTime)(this["CurrentDatabaseLastRefreshedOn"]));
            }
            set {
                this["CurrentDatabaseLastRefreshedOn"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("False")]
        public bool CurrentDatabaseIsPopulated {
            get {
                return ((bool)(this["CurrentDatabaseIsPopulated"]));
            }
            set {
                this["CurrentDatabaseIsPopulated"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("False")]
        public bool CurrentMappingIsPopulated {
            get {
                return ((bool)(this["CurrentMappingIsPopulated"]));
            }
            set {
                this["CurrentMappingIsPopulated"] = value;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("/sites/servicecatalogue")]
        public string URLSharePointSitePortionDEV {
            get {
                return ((string)(this["URLSharePointSitePortionDEV"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("/sites/uat-servicecatalogue")]
        public string URLSharePointSitePortionQA {
            get {
                return ((string)(this["URLSharePointSitePortionQA"]));
            }
        }
    }
}
