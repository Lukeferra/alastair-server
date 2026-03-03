using System;
using System.IO;
using Rhino;
using Rhino.Commands;
using Rhino.Input;
using Rhino.Input.Custom;
using ALASTAIR.Core;

namespace ALASTAIR.Commands
{
    /// <summary>
    /// ALASTAIR_Run — dispatcher command.
    ///
    /// Usage dalla macro toolbar:
    ///   ! _ALASTAIR_Run Id CreaTesto _Enter
    ///
    /// Il comando legge il parametro Id dalla command line,
    /// trova lo script corrispondente nel manifest locale ed lo esegue.
    /// </summary>
    [CommandStyle(Style.Hidden)]
    public class ALASTAIR_RunCommand : Command
    {
        public override string EnglishName => "ALASTAIR_Run";

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            // Legge il parametro Id come stringa dalla command line
            string commandId = string.Empty;

            var gs = new GetString();
            gs.SetCommandPrompt("ALASTAIR: Id comando");
            gs.AcceptNothing(false);

            // In modalità scripted (chiamata da macro) il valore è già sulla command line
            var getResult = gs.Get();

            if (getResult == GetResult.String)
            {
                commandId = gs.StringResult().Trim();
            }
            else
            {
                RhinoApp.WriteLine("ALASTAIR: Nessun Id specificato.");
                return Result.Nothing;
            }

            if (string.IsNullOrWhiteSpace(commandId))
            {
                RhinoApp.WriteLine("ALASTAIR: Id comando vuoto.");
                return Result.Nothing;
            }

            return ExecuteCommand(doc, commandId);
        }

        private Result ExecuteCommand(RhinoDoc doc, string commandId)
        {
            Logger.Log($"[ALASTAIR_Run] Esecuzione comando: {commandId}");

            // 1. Carica manifest locale
            Manifest manifest = CacheManager.LoadLocalManifest();
            if (manifest == null)
            {
                RhinoApp.WriteLine("ALASTAIR: Manifest locale non trovato. Esegui ALASTAIR_CheckUpdates.");
                Logger.Log("[ALASTAIR_Run] Manifest locale non trovato.");
                return Result.Failure;
            }

            // 2. Trova la voce corrispondente all'Id
            CommandEntry entry = manifest.Commands.Find(c =>
                string.Equals(c.Id, commandId, StringComparison.OrdinalIgnoreCase));

            if (entry == null)
            {
                RhinoApp.WriteLine($"ALASTAIR: Comando '{commandId}' non trovato nel manifest.");
                Logger.Log($"[ALASTAIR_Run] Comando '{commandId}' non trovato.");
                return Result.Failure;
            }

            // 3. Risolve il percorso dello script
            string scriptFile = entry.GetScriptForCurrentRhino();
            if (string.IsNullOrEmpty(scriptFile))
            {
                RhinoApp.WriteLine($"ALASTAIR: Nessuno script definito per '{commandId}'.");
                Logger.Log($"[ALASTAIR_Run] Nessun filename script per '{commandId}'.");
                return Result.Failure;
            }

            string scriptPath = CacheManager.GetScriptPath(scriptFile);
            if (!File.Exists(scriptPath))
            {
                RhinoApp.WriteLine($"ALASTAIR: Script non trovato in cache: {scriptFile}. Esegui ALASTAIR_CheckUpdates.");
                Logger.Log($"[ALASTAIR_Run] File script mancante: {scriptPath}");
                return Result.Failure;
            }

            // 4. Esegue lo script tramite _-RunPythonScript
            try
            {
                // Copia lo script in un file temp con nome semplice (evita problemi con percorsi lunghi)
                string tempScript = CacheManager.GetTempPath("_run_current.py");
                File.Copy(scriptPath, tempScript, overwrite: true);

                bool ok = RhinoApp.RunScript($"_-RunPythonScript \"{tempScript}\"", false);

                // Pulizia file temp
                try { File.Delete(tempScript); } catch { }

                if (!ok)
                {
                    Logger.Log($"[ALASTAIR_Run] RunPythonScript ha restituito false per '{scriptFile}'.");
                    return Result.Failure;
                }

                Logger.Log($"[ALASTAIR_Run] Script '{scriptFile}' eseguito con successo.");
                return Result.Success;
            }
            catch (Exception ex)
            {
                string msg = $"Errore script '{scriptFile}': {ex.Message}";
                RhinoApp.WriteLine($"ALASTAIR: {msg}");
                Logger.Log($"[ALASTAIR_Run ERROR] {msg}\n{ex.StackTrace}");
                return Result.Failure;
            }
        }
    }
}
