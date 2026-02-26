using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using Microsoft.Win32;

namespace InstallerActions
{
    [RunInstaller(true)]
    public class RegistryRemediation : Installer
    {
        private const string AddInProgId = "ZeroD.PhishReporter";
        private const string OfficeVersion = "16.0";

        private const string AddinsKeyPath =
            @"Software\Microsoft\Office\Outlook\Addins\" + AddInProgId;

        private const string ResiliencyBasePath =
            @"Software\Microsoft\Office\" + OfficeVersion + @"\Outlook\Resiliency";

        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);
            RemediateDisabledState();
        }

        public override void Commit(IDictionary savedState)
        {
            base.Commit(savedState);
        }

        public override void Rollback(IDictionary savedState)
        {
            base.Rollback(savedState);
        }

        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);
        }

        private void RemediateDisabledState()
        {
            try
            {
                // DEPL-02: Clear DisabledItems FIRST (before LoadBehavior reset)
                // Per Pitfall 2: if DisabledItems is not cleared, Outlook will
                // re-disable the add-in and reset LoadBehavior to 2 on next launch.
                ClearDisabledItems();

                // Clear CrashingAddinList as defensive measure
                // (add-in may have crashed pre-Phase 5 fixes)
                ClearCrashingAddinList();

                // DEPL-01: Reset LoadBehavior to 3 (load at startup)
                ResetLoadBehavior();
            }
            catch (Exception)
            {
                // Custom action must not fail the install.
                // Registry remediation is best-effort — if it fails,
                // DoNotDisableAddinList and AddinList keys provide fallback protection.
            }
        }

        private void ResetLoadBehavior()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(AddinsKeyPath, writable: true))
            {
                if (key == null) return; // Fresh install, Addins key written later by static entries

                var currentValue = key.GetValue("LoadBehavior");
                if (currentValue is int intValue && intValue != 3)
                {
                    key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                }
            }
        }

        private void ClearDisabledItems()
        {
            DeleteResiliencySubKey("DisabledItems");
        }

        private void ClearCrashingAddinList()
        {
            DeleteResiliencySubKey("CrashingAddinList");
        }

        private void DeleteResiliencySubKey(string subKeyName)
        {
            try
            {
                using (var resiliencyKey = Registry.CurrentUser.OpenSubKey(
                    ResiliencyBasePath, writable: true))
                {
                    if (resiliencyKey == null) return; // No Resiliency key on fresh machine

                    // Verify subkey exists before attempting delete
                    using (var targetKey = resiliencyKey.OpenSubKey(subKeyName))
                    {
                        if (targetKey == null) return;
                    }

                    resiliencyKey.DeleteSubKey(subKeyName, throwOnMissingSubKey: false);
                }
            }
            catch (Exception)
            {
                // Best-effort: if cleanup fails, DoNotDisableAddinList
                // and AddinList keys provide fallback protection.
            }
        }
    }
}
