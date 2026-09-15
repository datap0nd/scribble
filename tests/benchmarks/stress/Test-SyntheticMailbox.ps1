#requires -Version 5.1
param(
    [string]$AssemblyPath = (Join-Path $PSScriptRoot '../../../src/Scribble/bin/Release/Scribble.dll'),
    [string]$KitPath = (Join-Path $PSScriptRoot '../generated/stress-corpus'),
    [switch]$CompileOnly
)
$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') { throw 'Run this operator validation with powershell.exe -STA.' }
$assembly = (Resolve-Path $AssemblyPath).Path
Add-Type -TypeDefinition 'using System;using System.IO;using System.Reflection;public static class NativeMailboxResolver{public static void Install(string dir){AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(dir,new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};}}'
[NativeMailboxResolver]::Install((Split-Path $assembly))
Add-Type -Path $assembly
Add-Type -AssemblyName System.Windows.Forms,System.Web.Extensions
if ($null -ne [Scribble.Testing.TestLab]::Status() -or $null -ne [Scribble.Testing.TestLabSuite]::Active()) { throw 'A Test Lab session is active. Finish it before native corpus validation.' }
if ([Scribble.Testing.TestLabSuiteWindow]::HasLiveWindow(-1)) { throw 'Close the Test Lab window before native corpus validation.' }

# This probe submits only native query parameters to the production mailbox
# tool host. Expected IDs are compared afterwards, outside the tool host.
# The WinForms pump keeps every Outlook continuation on the original STA.
$probeSource = @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Scribble.Chat;
using Scribble.Outlook;
using Scribble.Testing;
public static class NativeMailboxProbe
{
    public static void Run(object application, string catalog, string output)
    {
        Exception failure = null;
        using (var window = new Form { ShowInTaskbar = false, Opacity = 0, Width = 1, Height = 1 })
        {
            window.Shown += async (sender, args) => {
                try { await Check(application, catalog, output); }
                catch (Exception error) { failure = error; }
                finally { window.Close(); }
            };
            Application.Run(window);
        }
        if (failure != null) throw new InvalidOperationException("Native mailbox validation failed.", failure);
    }
    private static async Task Check(object application, string catalog, string output)
    {
        var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        var data = (Dictionary<string, object>)json.DeserializeObject(File.ReadAllText(Path.Combine(catalog, "evaluator-only", "mail_catalog.json")));
        var tests = (object[])data["search_tasks"];
        var results = new List<object>();
        var started = DateTime.UtcNow;
        foreach (Dictionary<string, object> test in tests)
        {
            var id = Convert.ToString(test["case_id"]);
            var scope = (Dictionary<string, object>)test["search_scope"];
            var wanted = new HashSet<string>(((object[])test["expected_ids"]).Select(Convert.ToString), StringComparer.Ordinal);
            var observed = new HashSet<string>(StringComparer.Ordinal);
            var pages = new List<object>();
            var handles = new Dictionary<string, string>(StringComparer.Ordinal);
            var bodyChecks = new List<object>();
            using (var host = new MailboxToolHost(application, null))
            {
                string cursor = "";
                for (int page = 1; ; page++)
                {
                    if (page > 1000) throw new InvalidDataException("Native cursor did not terminate: " + id);
                    var arguments = new Dictionary<string, object>(scope);
                    arguments["cursor"] = cursor;
                    var call = new ChatToolCall { id = id + "-" + page, type = "function", function = new ChatToolCallFunction { name = "search_mailbox", arguments = json.Serialize(arguments) } };
                    var result = await host.ExecuteAsync(call, CancellationToken.None);
                    var response = (Dictionary<string, object>)json.DeserializeObject(result.Content);
                    if (response.ContainsKey("error_code")) throw new InvalidDataException(id + ": " + result.Content);
                    pages.Add(response);
                    foreach (Dictionary<string, object> message in (object[])response["results"])
                    {
                        var match = Regex.Match(Convert.ToString(message["subject"]), @"^\[(MAIL[0-9]{4})\]");
                        if (!match.Success || !observed.Add(match.Groups[1].Value)) throw new InvalidDataException(id + ": missing or duplicate native fixture identity.");
                        handles[match.Groups[1].Value] = Convert.ToString(message["handle"]);
                    }
                    if (Convert.ToBoolean(response["enumeration_complete"])) break;
                    cursor = Convert.ToString(response["next_cursor"]);
                    if (cursor.Length == 0) throw new InvalidDataException(id + ": incomplete enumeration without a cursor.");
                }
                if (id == "OL65")
                {
                    foreach (var found in handles)
                    {
                        var body = new System.Text.StringBuilder();
                        int offset = 0, bodyPages = 0;
                        for (;;)
                        {
                            var readArguments = new Dictionary<string, object> { { "handles", new[] { found.Value } }, { "body_offset", offset } };
                            var call = new ChatToolCall { id = id + "-body-" + found.Key + "-" + offset, type = "function", function = new ChatToolCallFunction { name = "read_messages", arguments = json.Serialize(readArguments) } };
                            var response = (Dictionary<string, object>)json.DeserializeObject((await host.ExecuteAsync(call, CancellationToken.None)).Content);
                            if (response.ContainsKey("error_code")) throw new InvalidDataException(id + ": body read failed.");
                            var message = (Dictionary<string, object>)((object[])response["messages"])[0];
                            body.Append(Convert.ToString(message["body"])); bodyPages++;
                            if (Convert.ToBoolean(message["body_complete"])) break;
                            var next = Convert.ToInt32(message["next_body_offset"]);
                            if (next <= offset || bodyPages > 100) throw new InvalidDataException("Body pagination did not advance.");
                            offset = next;
                        }
                        var marker = found.Key + "-TAIL";
                        if (bodyPages < 2 || !body.ToString().Contains(marker)) throw new InvalidDataException("Long native body was truncated before its tail: " + found.Key);
                        bodyChecks.Add(new { id = found.Key, pages = bodyPages, tail_marker = marker, passed = true });
                    }
                }
            }
            var passed = wanted.SetEquals(observed);
            var entry = new { case_id = id, passed, expected_count = wanted.Count, actual_count = observed.Count,
                missing = wanted.Except(observed).ToArray(), unexpected = observed.Except(wanted).ToArray(), pages, body_checks = bodyChecks };
            results.Add(entry);
            File.WriteAllText(Path.Combine(output, id + "-native-search.json"), json.Serialize(entry));
            Console.WriteLine(id + ": " + observed.Count + " IDs across " + pages.Count + " native pages — " + (passed ? "PASS" : "FAIL"));
            if (!passed) throw new InvalidDataException("Native scope disagrees with the independent source oracle: " + id);
        }
        // An arbitrary foreign identity must be rejected before Outlook opens it.
        bool refused = false;
        try { TestLabMailbox.ValidateIdentity("not-a-fixture-entry", "not-the-fixture-store"); } catch (InvalidOperationException) { refused = true; }
        if (!refused) throw new InvalidOperationException("Foreign mailbox identity was not rejected.");
        File.WriteAllText(Path.Combine(output, "validation.json"), json.Serialize(new {
            status = "passed", kind = "native_adapter_preflight_not_model_evaluation", started_utc = started, finished_utc = DateTime.UtcNow,
            tested_searches = tests.Length, fixture_messages = 500, foreign_identity_rejected = refused, results }));
    }
}
'@
Add-Type -TypeDefinition $probeSource -ReferencedAssemblies $assembly,'System.Core.dll','System.Windows.Forms.dll','System.Web.Extensions.dll'
if ($CompileOnly) { Write-Host 'Native mailbox probe compiled; no Office or model calls made.'; return }
$source = (Resolve-Path $KitPath).Path
$folder = Join-Path (Split-Path $source) ('native-mailbox-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $folder | Out-Null
$state = New-Object Scribble.Testing.SuiteState
$state.id = [guid]::NewGuid().ToString('N'); $state.folder = $folder
$state.catalogRoot = Join-Path $folder 'catalog'; $state.fixtureSuiteId = 'scribble-stress-v1'
$state.caseId = 'OL01'; $state.host = 'Outlook'; $state.pid = $PID
$state.processStart = [Diagnostics.Process]::GetCurrentProcess().StartTime.ToUniversalTime().Ticks
$state.expires = [DateTime]::UtcNow.AddHours(3)
$outlook = $null
$filter = $null
try {
    $hash = [Scribble.Testing.TestLab]::FileHash((Join-Path $source 'manifest.json'))
    $state.kitHash = $hash
    [void][Scribble.Testing.TestLabSuite]::SnapshotExternalKit($source, $hash, $state.catalogRoot, [Threading.CancellationToken]::None)
    [Scribble.Testing.TestLabSuite]::Save($state)
    [Scribble.Testing.TestLab]::Enable($state.catalogRoot)
    $filter = New-Object Scribble.Testing.TestLabComMessageFilter([Threading.CancellationToken]::None)
    $outlook = New-Object -ComObject Outlook.Application
    [Scribble.Testing.TestLabMailbox]::Prepare($outlook, [Threading.CancellationToken]::None, [Action[string]]{param($message) Write-Host $message})
    $run = [Scribble.Testing.TestLab]::Start('OL01', 'Outlook', $true)
    $state.runId = $run.run_id; [Scribble.Testing.TestLabSuite]::Save($state)
    [NativeMailboxProbe]::Run($outlook, $state.catalogRoot, $folder)
    [Scribble.Testing.TestLab]::Finish($true)
    Write-Host ('Native 70-query preflight passed. Evidence: ' + (Join-Path $folder 'validation.json'))
} finally {
    if ($null -ne [Scribble.Testing.TestLab]::Status()) { [Scribble.Testing.TestLab]::Disable() }
    $state.expires = [DateTime]::UtcNow.AddMinutes(-1); [Scribble.Testing.TestLabSuite]::Save($state)
    if ($null -ne $filter) { $filter.Dispose() }
    if ($null -ne $outlook -and [Runtime.InteropServices.Marshal]::IsComObject($outlook)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) }
}
