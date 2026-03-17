"""HydroEchoToGnssVerticalOffset

Trimble Business Center (TBC) macro.

Purpose
-------
Convert Trimble Access Hydro JobXML (.jxl) EchoSounding depth observations into
GNSSVerticalOffset observations so the file can be imported into TBC as a GNSS
vertical offset.

What it does
------------
- Finds PointRecord elements that contain an <EchoSounding> element.
- Reads <Depth> (or optionally <SecondaryDepth>)
- Creates/replaces a <GNSSVerticalOffset><VerticalOffset>...</VerticalOffset></GNSSVerticalOffset>
  using the NEGATIVE of the depth (Depth is assumed positive down).
- Removes the <EchoSounding> element (and optionally EchoSounderConfigurationRecord blocks)
  to improve compatibility with older TBC JobXML importers.
- Optionally stores the removed EchoSounding values into point Notes.
- Saves a new JXL file and optionally imports it into the current project.

Notes
-----
- GNSSVerticalOffset values in JobXML are stored in meters.
- EchoSounding Depth values are assumed to be meters (consistent with other GNSS
  observation values in JobXML). If your depths appear to be in US survey feet,
  enable the "Convert depth from US survey feet to meters" option.

Version
-------
1.2

Fixes in 1.2
------------
- Keeps the working conversion logic unchanged.
- Replaces hardcoded import item type imports with reflection-based discovery so
  the import step works across more TBC builds.
- Fixes the button caption to show a single ampersand.
- Adds footer text requested by the user.
"""

import clr


def _try_add_reference(name):
    try:
        clr.AddReference(name)
        return True
    except:
        return False


# IronPython 3 in TBC often needs explicit assembly loads.
for _asm in [
    'System',
    'System.Core',
    'System.Xml',
    'PresentationCore',
    'PresentationFramework',
    'WindowsBase',
    'IronPython.Wpf',
]:
    _try_add_reference(_asm)

import wpf

from System import Activator, AppDomain, Array, DateTime, Double, String, Enum, Type
from System.Globalization import CultureInfo
from System.IO import Directory, File, Path, StreamReader
from System.Reflection import BindingFlags
from System.Text import Encoding
from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult, Window
from System.Xml import XmlDocument
from Microsoft.Win32 import OpenFileDialog, SaveFileDialog


# -------------------------------
# TBC command registration
# -------------------------------

def Setup(cmdData, macroFileFolder=None):
    """Called once when TBC scans the MacroCommands folder."""
    try:
        cmdData.Key = "HydroEchoToGnssVerticalOffset"
    except:
        pass
    try:
        cmdData.CommandName = "HydroEchoToGnssVerticalOffset"
    except:
        pass
    try:
        cmdData.Caption = "Hydro: EchoSounding -> GNSS Vertical Offset"
    except:
        pass

    for propName in ["ToolTip", "Tooltip"]:
        try:
            setattr(cmdData, propName, "Convert Hydro JobXML EchoSounding depths into GNSS Vertical Offsets, save, and import.")
        except:
            pass

    try:
        if macroFileFolder:
            _try_add_reference('System.Drawing')
            from System.Drawing import Bitmap
            bmpPath = Path.Combine(macroFileFolder, cmdData.Key + ".bmp")
            if File.Exists(bmpPath):
                b = Bitmap(bmpPath)
                cmdData.ImageSmall = b
                cmdData.ImageLarge = b
    except:
        pass


def CanExecute(cmd, currentProject, parameters):
    return True



def Execute(cmd, currentProject, macroFileFolder, parameters):
    try:
        dlg = HydroEchoToGnssVerticalOffsetWindow(currentProject, macroFileFolder, cmd)
        dlg.ShowDialog()
    except Exception as ex:
        MessageBox.Show(
            "Unexpected error launching macro:\n\n{0}".format(str(ex)),
            "Hydro: EchoSounding -> GNSS Vertical Offset",
            MessageBoxButton.OK,
            MessageBoxImage.Error,
        )


# -------------------------------
# UI Window
# -------------------------------

class HydroEchoToGnssVerticalOffsetWindow(Window):
    def __init__(self, currentProject, macroFileFolder, cmd=None):
        self._currentProject = currentProject
        self._macroFileFolder = macroFileFolder
        self._cmd = cmd
        self._invariant = CultureInfo.InvariantCulture
        self._service_instance_cache = {}

        xamlPath = Path.Combine(macroFileFolder, "HydroEchoToGnssVerticalOffset.xaml")
        sr = None
        try:
            sr = StreamReader(xamlPath)
            wpf.LoadComponent(self, sr)
        finally:
            try:
                if sr is not None:
                    sr.Close()
            except:
                pass

        self.btnBrowseInput.Click += self._browse_input
        self.btnBrowseOutput.Click += self._browse_output
        self.btnRun.Click += self._run
        self.btnClose.Click += self._close
        self.chkImportIntoTbc.Click += self._refresh_run_button_text

        self.chkRemoveEchoSounding.IsChecked = True
        self.chkRemoveEchoConfig.IsChecked = True
        self.chkUseSecondaryDepth.IsChecked = True
        self.chkKeepExtraAsNotes.IsChecked = True
        self.chkConvertDepthFromFeet.IsChecked = False
        self.chkOverwriteExisting.IsChecked = False
        self.chkImportIntoTbc.IsChecked = True

        self._refresh_run_button_text(None, None)
        self._log("Ready. Select a .JXL file to convert.")

    # ---------------------------
    # UI helpers
    # ---------------------------

    def _refresh_run_button_text(self, sender, args):
        try:
            if self.chkImportIntoTbc.IsChecked:
                self.btnRun.Content = "Convert & Import"
            else:
                self.btnRun.Content = "Convert"
        except:
            pass

    def _clear_log(self):
        try:
            self.txtLog.Text = ""
        except:
            pass

    def _log(self, msg):
        try:
            stamp = DateTime.Now.ToString("HH:mm:ss")
            self.txtLog.AppendText("[{0}] {1}\r\n".format(stamp, msg))
            self.txtLog.ScrollToEnd()
        except:
            pass

    def _show_error(self, msg, details=None):
        full = msg
        if details:
            full = msg + "\n\nDetails:\n" + details
        MessageBox.Show(full, "Hydro: EchoSounding -> GNSS Vertical Offset", MessageBoxButton.OK, MessageBoxImage.Error)

    def _show_info(self, msg):
        MessageBox.Show(msg, "Hydro: EchoSounding -> GNSS Vertical Offset", MessageBoxButton.OK, MessageBoxImage.Information)

    def _bool(self, checkbox):
        try:
            return bool(checkbox.IsChecked)
        except:
            return False

    def _to_invariant_double_text(self, value):
        return Double(value).ToString("G17", self._invariant)

    # ---------------------------
    # Browse buttons
    # ---------------------------

    def _browse_input(self, sender, args):
        dlg = OpenFileDialog()
        dlg.Filter = "JobXML (*.jxl)|*.jxl|All files (*.*)|*.*"
        dlg.Multiselect = False

        if dlg.ShowDialog() == True:
            self.txtInputPath.Text = dlg.FileName
            try:
                folder = Path.GetDirectoryName(dlg.FileName)
                baseName = Path.GetFileNameWithoutExtension(dlg.FileName)
                outName = baseName + "_GNSSVertOffset.jxl"
                self.txtOutputPath.Text = Path.Combine(folder, outName)
            except:
                pass

    def _browse_output(self, sender, args):
        dlg = SaveFileDialog()
        dlg.Filter = "JobXML (*.jxl)|*.jxl|All files (*.*)|*.*"

        try:
            if self.txtOutputPath.Text and self.txtOutputPath.Text.strip() != "":
                dlg.FileName = self.txtOutputPath.Text
        except:
            pass

        if dlg.ShowDialog() == True:
            self.txtOutputPath.Text = dlg.FileName

    # ---------------------------
    # Run
    # ---------------------------

    def _run(self, sender, args):
        self.btnRun.IsEnabled = False

        try:
            self._clear_log()

            inputPath = (self.txtInputPath.Text or "").strip()
            outputPath = (self.txtOutputPath.Text or "").strip()

            if inputPath == "":
                self._show_error("Please select an input .JXL file.")
                return

            if not File.Exists(inputPath):
                self._show_error("Input file does not exist:", inputPath)
                return

            if outputPath == "":
                self._show_error("Please specify an output .JXL file.")
                return

            if str(Path.GetFullPath(inputPath)).lower() == str(Path.GetFullPath(outputPath)).lower():
                self._show_error("Output path must be different from input path (to avoid overwriting).")
                return

            if File.Exists(outputPath):
                r = MessageBox.Show(
                    "The output file already exists.\n\nOverwrite?\n\n{0}".format(outputPath),
                    "Hydro: EchoSounding -> GNSS Vertical Offset",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question,
                )
                if r != MessageBoxResult.Yes:
                    self._log("Cancelled (output file exists).")
                    return

            options = {
                "removeEchoSounding": self._bool(self.chkRemoveEchoSounding),
                "removeEchoConfig": self._bool(self.chkRemoveEchoConfig),
                "useSecondaryDepth": self._bool(self.chkUseSecondaryDepth),
                "keepExtraAsNotes": self._bool(self.chkKeepExtraAsNotes),
                "convertDepthFromFeet": self._bool(self.chkConvertDepthFromFeet),
                "overwriteExisting": self._bool(self.chkOverwriteExisting),
                "importIntoTbc": self._bool(self.chkImportIntoTbc),
            }

            self._log("Input : {0}".format(inputPath))
            self._log("Output: {0}".format(outputPath))
            self._log("Loading XML...")

            report = self._convert_file(inputPath, outputPath, options)

            reportPath = self._write_report_file(outputPath, report)
            if reportPath:
                self._log("Report file: {0}".format(reportPath))

            self._log("---")
            for line in report.split("\n"):
                if line.strip() != "":
                    self._log(line)

            if options["importIntoTbc"]:
                self._log("---")
                self._log("Importing converted JXL into current TBC project...")
                ok, err = self._import_into_tbc(outputPath)
                if ok:
                    self._log("Import complete.")
                    self._show_info("Conversion complete and file imported successfully.\n\nOutput:\n{0}".format(outputPath))
                else:
                    self._log("IMPORT FAILED: {0}".format(err))
                    self._show_error(
                        "Conversion complete, but TBC import failed.\n\nYou can still import the output file manually from File > Import.",
                        "Output file:\n{0}\n\nImport error:\n{1}".format(outputPath, err),
                    )
            else:
                self._show_info("Conversion complete.\n\nOutput:\n{0}".format(outputPath))

        except Exception as ex:
            self._log("ERROR: {0}".format(str(ex)))
            self._show_error("An unexpected error occurred.", str(ex))

        finally:
            self.btnRun.IsEnabled = True

    # ---------------------------
    # Conversion
    # ---------------------------

    def _convert_file(self, inputPath, outputPath, options):
        doc = XmlDocument()
        doc.PreserveWhitespace = True
        doc.Load(inputPath)

        removedEchoConfig = 0
        if options["removeEchoConfig"]:
            cfgNodesLive = doc.GetElementsByTagName("EchoSounderConfigurationRecord")
            cfgNodes = [n for n in cfgNodesLive]
            for n in cfgNodes:
                if n.ParentNode is not None:
                    n.ParentNode.RemoveChild(n)
                    removedEchoConfig += 1

        pointNodesLive = doc.GetElementsByTagName("PointRecord")
        pointNodes = [p for p in pointNodesLive]

        totalPoints = len(pointNodes)
        echoPoints = 0
        converted = 0
        updatedExisting = 0
        skippedNoDepth = 0
        skippedBadDepth = 0
        notesAdded = 0
        errorSamples = []

        for p in pointNodes:
            echo = p.SelectSingleNode("./EchoSounding")
            if echo is None:
                continue

            echoPoints += 1

            depthText, depthSource = self._get_depth_text(echo, options["useSecondaryDepth"])
            if depthText == "":
                skippedNoDepth += 1
                if options["removeEchoSounding"]:
                    try:
                        p.RemoveChild(echo)
                    except:
                        pass
                continue

            try:
                depthVal = self._parse_double(depthText)
            except Exception:
                skippedBadDepth += 1
                if len(errorSamples) < 25:
                    errorSamples.append("Point {0}: invalid {1} '{2}'".format(self._get_point_name(p), depthSource, depthText))
                if options["removeEchoSounding"]:
                    try:
                        p.RemoveChild(echo)
                    except:
                        pass
                continue

            if options["convertDepthFromFeet"]:
                depthVal = float(depthVal) * 0.3048006096012192

            vertOffset = -float(depthVal)
            vertText = self._to_invariant_double_text(vertOffset)

            noteText = None
            if options["keepExtraAsNotes"]:
                noteText = self._build_echo_note(echo, depthText, depthSource)

            existingGnss = p.SelectSingleNode("./GNSSVerticalOffset")
            if existingGnss is not None:
                if options["overwriteExisting"]:
                    vo = existingGnss.SelectSingleNode("./VerticalOffset")
                    if vo is None:
                        vo = doc.CreateElement("VerticalOffset")
                        existingGnss.AppendChild(vo)
                    vo.InnerText = vertText
                    updatedExisting += 1

                if options["removeEchoSounding"]:
                    try:
                        p.RemoveChild(echo)
                    except:
                        pass
            else:
                gnssNode = doc.CreateElement("GNSSVerticalOffset")
                voNode = doc.CreateElement("VerticalOffset")
                voNode.InnerText = vertText
                gnssNode.AppendChild(voNode)

                if options["removeEchoSounding"]:
                    try:
                        p.ReplaceChild(gnssNode, echo)
                    except:
                        p.AppendChild(gnssNode)
                        try:
                            p.RemoveChild(echo)
                        except:
                            pass
                else:
                    try:
                        p.InsertAfter(gnssNode, echo)
                    except:
                        p.AppendChild(gnssNode)

                converted += 1

            if noteText:
                try:
                    if self._add_note(doc, p, noteText):
                        notesAdded += 1
                except:
                    pass

        outFolder = Path.GetDirectoryName(outputPath)
        if outFolder and not Directory.Exists(outFolder):
            Directory.CreateDirectory(outFolder)

        doc.Save(outputPath)

        lines = []
        lines.append("Conversion summary")
        lines.append("  Total PointRecord nodes             : {0}".format(totalPoints))
        lines.append("  Points containing EchoSounding      : {0}".format(echoPoints))
        lines.append("  New GNSSVerticalOffset created      : {0}".format(converted))
        lines.append("  Existing GNSSVerticalOffset updated : {0}".format(updatedExisting))
        lines.append("  Skipped (missing Depth)             : {0}".format(skippedNoDepth))
        lines.append("  Skipped (invalid Depth)             : {0}".format(skippedBadDepth))
        if options["removeEchoConfig"]:
            lines.append("  EchoSounderConfigurationRecord removed : {0}".format(removedEchoConfig))
        lines.append("  Notes added                         : {0}".format(notesAdded))

        if len(errorSamples) > 0:
            lines.append("")
            lines.append("Sample issues (first {0}):".format(len(errorSamples)))
            for e in errorSamples:
                lines.append("  - " + e)

        return "\n".join(lines)

    def _write_report_file(self, outputJxlPath, reportText):
        try:
            folder = Path.GetDirectoryName(outputJxlPath)
            baseName = Path.GetFileNameWithoutExtension(outputJxlPath)
            reportPath = Path.Combine(folder, baseName + "_conversion_report.txt")
            File.WriteAllText(reportPath, reportText, Encoding.UTF8)
            return reportPath
        except:
            return None

    def _get_depth_text(self, echoNode, useSecondary):
        depthSource = "Depth"
        depthText = ""
        d = echoNode.SelectSingleNode("./Depth")
        if d is not None and d.InnerText:
            depthText = d.InnerText.strip()

        if depthText == "" and useSecondary:
            depthSource = "SecondaryDepth"
            s = echoNode.SelectSingleNode("./SecondaryDepth")
            if s is not None and s.InnerText:
                depthText = s.InnerText.strip()

        return depthText, depthSource

    def _parse_double(self, s):
        return Double.Parse(s, self._invariant)

    def _get_point_name(self, pointNode):
        try:
            n = pointNode.SelectSingleNode("./Name")
            if n is not None and n.InnerText:
                return n.InnerText.strip()
        except:
            pass
        try:
            if pointNode.Attributes and pointNode.Attributes["ID"]:
                return "ID=" + pointNode.Attributes["ID"].Value
        except:
            pass
        return "(unknown)"

    def _build_echo_note(self, echoNode, depthText, depthSource):
        parts = []
        parts.append("EchoSounding")
        if depthText != "":
            parts.append("{0}={1}".format(depthSource, depthText))

        try:
            sd = echoNode.SelectSingleNode("./SecondaryDepth")
            if sd is not None and sd.InnerText and sd.InnerText.strip() != "":
                parts.append("SecondaryDepth={0}".format(sd.InnerText.strip()))
        except:
            pass

        try:
            extra = echoNode.SelectSingleNode("./ExtraData")
            if extra is not None:
                dataNodes = extra.SelectNodes("./Data")
                for dn in dataNodes:
                    try:
                        nameNode = dn.SelectSingleNode("./Name")
                        valNode = dn.SelectSingleNode("./Value")
                        if nameNode is None or not nameNode.InnerText:
                            continue
                        name = nameNode.InnerText.strip()
                        val = valNode.InnerText.strip() if (valNode is not None and valNode.InnerText) else ""
                        if val != "":
                            parts.append("{0}={1}".format(name, val))
                        else:
                            parts.append("{0}".format(name))
                    except:
                        pass
        except:
            pass

        return "; ".join(parts)

    def _add_note(self, doc, pointNode, noteText):
        if noteText is None or noteText.strip() == "":
            return False

        notes = pointNode.SelectSingleNode("./Notes")
        if notes is None:
            notes = doc.CreateElement("Notes")
            afterNode = pointNode.SelectSingleNode("./GNSSVerticalOffset")
            if afterNode is not None:
                try:
                    pointNode.InsertAfter(notes, afterNode)
                except:
                    pointNode.AppendChild(notes)
            else:
                pointNode.AppendChild(notes)

        note = doc.CreateElement("Note")
        note.InnerText = noteText
        notes.AppendChild(note)
        return True

    # ---------------------------
    # Import into TBC
    # ---------------------------

    def _import_into_tbc(self, jxlPath):
        load_attempts = []
        for asm_name in [
            "Trimble.Vce.UI.BaseCommands",
            "Trimble.Vce.BaseCommands",
            "Trimble.Vce.UI.Commands",
            "Trimble.Sdk.UI",
        ]:
            try:
                if _try_add_reference(asm_name):
                    load_attempts.append("loaded " + asm_name)
            except:
                pass

        try:
            import_service_type = self._find_import_service_type()
            pieces = []

            if import_service_type is not None:
                ok, detail = self._try_invoke_import_service(import_service_type, jxlPath)
                if ok:
                    return True, None
                if detail:
                    pieces.append(detail)
            else:
                pieces.append("ImportService type was not found in loaded TBC assemblies.")

            ok2, detail2 = self._try_project_object_import(jxlPath)
            if ok2:
                return True, None
            if detail2:
                pieces.append(detail2)

            if len(pieces) > 0:
                return False, "; ".join(pieces)
            return False, "No usable import method was found on ImportService or currentProject."
        except Exception as ex:
            prefix = "; ".join(load_attempts)
            if prefix:
                return False, prefix + "; " + str(ex)
            return False, str(ex)

    def _find_import_service_type(self):
        candidates = []
        for asm in AppDomain.CurrentDomain.GetAssemblies():
            try:
                for t in asm.GetTypes():
                    try:
                        name = t.Name or ""
                        full = t.FullName or ""
                        if name == "ImportService" or full.endswith(".ImportService"):
                            candidates.append(t)
                    except:
                        pass
            except:
                pass

        preferred = []
        other = []
        for t in candidates:
            try:
                full = t.FullName or ""
                if full.startswith("Trimble."):
                    preferred.append(t)
                else:
                    other.append(t)
            except:
                other.append(t)

        if len(preferred) > 0:
            return preferred[0]
        if len(other) > 0:
            return other[0]
        return None

    def _try_invoke_import_service(self, importServiceType, jxlPath):
        errors = []
        method_names = ["ImportFiles", "ImportFile", "Import"]
        for method_name in method_names:
            methods = self._get_candidate_methods(importServiceType, method_name)
            for method in methods:
                try:
                    pars = method.GetParameters()
                    can_build, args = self._build_method_arguments(pars, jxlPath)
                    if not can_build:
                        continue
                    self._invoke_service_method(importServiceType, method, args)
                    return True, None
                except Exception as ex:
                    sig = self._format_method_signature(method)
                    errors.append(sig + ": " + self._unwrap_invoke_exception(ex))

        if len(errors) > 0:
            return False, "ImportService attempts failed: " + "; ".join(errors)
        return False, None

    def _get_candidate_methods(self, serviceType, methodName):
        results = []
        for m in serviceType.GetMethods():
            try:
                if m.Name == methodName:
                    results.append(m)
            except:
                pass
        results.sort(key=lambda m: len(m.GetParameters()))
        return results

    def _build_method_arguments(self, parameters, jxlPath):
        args = []
        used_file_arg = False
        for p in parameters:
            try:
                ptype = p.ParameterType
            except:
                return False, None

            if not used_file_arg:
                ok, value = self._try_build_file_argument(ptype, jxlPath)
                if ok:
                    args.append(value)
                    used_file_arg = True
                    continue

            ok, value = self._try_build_support_argument(p, jxlPath)
            if not ok:
                return False, None
            args.append(value)

        if not used_file_arg:
            return False, None
        return True, Array[object](args)

    def _try_build_file_argument(self, paramType, jxlPath):
        try:
            stringType = clr.GetClrType(String)
            if paramType == stringType:
                return True, jxlPath

            if paramType.IsArray:
                elementType = paramType.GetElementType()
                ok, single = self._try_build_file_argument(elementType, jxlPath)
                if ok:
                    arr = Array.CreateInstance(elementType, 1)
                    arr.SetValue(single, 0)
                    return True, arr

            if self._is_generic_sequence_type(paramType):
                elementType = paramType.GetGenericArguments()[0]
                ok, single = self._try_build_file_argument(elementType, jxlPath)
                if ok:
                    arr = Array.CreateInstance(elementType, 1)
                    arr.SetValue(single, 0)
                    return True, arr

            if self._is_tuple2_type(paramType):
                genericArgs = paramType.GetGenericArguments()
                leftType = genericArgs[0]
                rightType = genericArgs[1]
                ok, leftValue = self._try_build_file_argument(leftType, jxlPath)
                if not ok:
                    return False, None
                ok, rightValue = self._try_resolve_runtime_object(rightType, jxlPath)
                if not ok:
                    return False, None
                tupleValue = Activator.CreateInstance(paramType, Array[object]([leftValue, rightValue]))
                return True, tupleValue

            item = self._create_import_item_instance(paramType, jxlPath)
            return True, item
        except:
            return False, None

    def _try_build_support_argument(self, paramInfo, jxlPath):
        try:
            ptype = paramInfo.ParameterType
        except:
            return False, None

        try:
            if paramInfo.IsOptional:
                defaultValue = paramInfo.DefaultValue
                return True, defaultValue
        except:
            pass

        return self._try_resolve_runtime_object(ptype, jxlPath)

    def _try_resolve_runtime_object(self, targetType, jxlPath):
        try:
            if targetType is None:
                return False, None

            if targetType == clr.GetClrType(String):
                return True, ""

            try:
                if self._currentProject is not None and targetType.IsInstanceOfType(self._currentProject):
                    return True, self._currentProject
            except:
                pass

            try:
                if self._cmd is not None and targetType.IsInstanceOfType(self._cmd):
                    return True, self._cmd
            except:
                pass

            try:
                if targetType.IsEnum:
                    vals = Enum.GetValues(targetType)
                    if vals is not None and vals.Length > 0:
                        return True, vals.GetValue(0)
            except:
                pass

            try:
                full = targetType.FullName or ""
                if full == "System.Boolean":
                    return True, False
                if full == "System.Int32":
                    return True, 0
                if full == "System.Double":
                    return True, 0.0
            except:
                pass

            service = self._resolve_service_instance(targetType)
            if service is not None:
                return True, service

            if self._is_tuple2_type(targetType):
                args = targetType.GetGenericArguments()
                ok1, v1 = self._try_resolve_runtime_object(args[0], jxlPath)
                ok2, v2 = self._try_resolve_runtime_object(args[1], jxlPath)
                if ok1 and ok2:
                    return True, Activator.CreateInstance(targetType, Array[object]([v1, v2]))

            try:
                value = Activator.CreateInstance(targetType)
                return True, value
            except:
                pass
        except:
            pass

        return False, None

    def _is_generic_sequence_type(self, t):
        try:
            if not t.IsGenericType:
                return False
            full = t.GetGenericTypeDefinition().FullName or ""
            return (
                full.startswith("System.Collections.Generic.IEnumerable`1")
                or full.startswith("System.Collections.Generic.IList`1")
                or full.startswith("System.Collections.Generic.ICollection`1")
                or full.startswith("System.Collections.Generic.List`1")
            )
        except:
            return False

    def _is_tuple2_type(self, t):
        try:
            if not t.IsGenericType:
                return False
            full = t.GetGenericTypeDefinition().FullName or ""
            return full.startswith("System.Tuple`2") or full.startswith("System.ValueTuple`2")
        except:
            return False

    def _format_method_signature(self, method):
        try:
            names = []
            for p in method.GetParameters():
                try:
                    names.append(p.ParameterType.Name)
                except:
                    names.append("?")
            return method.Name + "(" + ", ".join(names) + ")"
        except:
            try:
                return method.Name
            except:
                return "(unknown method)"

    def _resolve_service_instance(self, serviceType):
        key = None
        try:
            key = serviceType.FullName
        except:
            key = str(serviceType)

        try:
            if key in self._service_instance_cache:
                return self._service_instance_cache[key]
        except:
            pass

        instance = self._resolve_service_instance_uncached(serviceType)
        try:
            self._service_instance_cache[key] = instance
        except:
            pass
        return instance

    def _resolve_service_instance_uncached(self, serviceType):
        try:
            if self._currentProject is not None and serviceType.IsInstanceOfType(self._currentProject):
                return self._currentProject
        except:
            pass
        try:
            if self._cmd is not None and serviceType.IsInstanceOfType(self._cmd):
                return self._cmd
        except:
            pass

        instance = self._try_static_singleton_on_type(serviceType)
        if instance is not None:
            return instance

        for root in [self._cmd, self._currentProject]:
            instance = self._search_object_for_service(root, serviceType, 0, {})
            if instance is not None:
                return instance

        instance = self._search_loaded_static_roots_for_service(serviceType)
        if instance is not None:
            return instance

        instance = self._try_construct_with_resolved_parameters(serviceType)
        if instance is not None:
            return instance

        return None

    def _try_static_singleton_on_type(self, serviceType):
        for name in ["Instance", "Current", "Default", "Singleton"]:
            try:
                prop = serviceType.GetProperty(name, BindingFlags.Public | BindingFlags.Static)
                if prop is not None:
                    value = prop.GetValue(None, None)
                    if value is not None:
                        return value
            except:
                pass
            try:
                field = serviceType.GetField(name, BindingFlags.Public | BindingFlags.Static)
                if field is not None:
                    value = field.GetValue(None)
                    if value is not None:
                        return value
            except:
                pass
        return None

    def _search_loaded_static_roots_for_service(self, serviceType):
        interesting_type_words = ["Service", "Locator", "Provider", "Application", "Shell", "Context", "Host", "Manager"]
        interesting_member_names = ["Instance", "Current", "Default", "Services", "ServiceProvider", "Provider", "Application", "App", "Shell", "Host", "Context"]

        for asm in AppDomain.CurrentDomain.GetAssemblies():
            try:
                for t in asm.GetTypes():
                    try:
                        type_name = t.Name or ""
                        full_name = t.FullName or ""
                        if not any(word in type_name or word in full_name for word in interesting_type_words):
                            continue

                        direct = None
                        try:
                            if t.IsClass and serviceType.IsAssignableFrom(t):
                                direct = self._try_static_singleton_on_type(t)
                        except:
                            pass
                        if direct is not None and serviceType.IsInstanceOfType(direct):
                            return direct

                        for member_name in interesting_member_names:
                            try:
                                prop = t.GetProperty(member_name, BindingFlags.Public | BindingFlags.Static)
                                if prop is not None:
                                    value = prop.GetValue(None, None)
                                    found = self._search_object_for_service(value, serviceType, 0, {})
                                    if found is not None:
                                        return found
                            except:
                                pass
                            try:
                                field = t.GetField(member_name, BindingFlags.Public | BindingFlags.Static)
                                if field is not None:
                                    value = field.GetValue(None)
                                    found = self._search_object_for_service(value, serviceType, 0, {})
                                    if found is not None:
                                        return found
                            except:
                                pass
                    except:
                        pass
            except:
                pass
        return None

    def _search_object_for_service(self, obj, serviceType, depth, visited):
        if obj is None:
            return None
        if depth > 3:
            return None

        try:
            obj_id = id(obj)
            if obj_id in visited:
                return None
            visited[obj_id] = True
        except:
            pass

        try:
            if serviceType.IsInstanceOfType(obj):
                return obj
        except:
            pass

        method_names = ["GetService", "Resolve", "TryGetService", "Locate"]
        for method_name in method_names:
            try:
                objType = obj.GetType()
                for method in objType.GetMethods():
                    try:
                        if method.Name != method_name:
                            continue
                        pars = method.GetParameters()
                        if len(pars) != 1:
                            continue
                        ptype = pars[0].ParameterType
                        value = None
                        if ptype == clr.GetClrType(Type):
                            value = method.Invoke(obj, Array[object]([serviceType]))
                        elif ptype == clr.GetClrType(String):
                            value = method.Invoke(obj, Array[object]([serviceType.FullName]))
                        else:
                            continue
                        if value is not None:
                            found = self._search_object_for_service(value, serviceType, depth + 1, visited)
                            if found is not None:
                                return found
                    except:
                        pass
            except:
                pass

        member_names = ["Services", "ServiceProvider", "Provider", "Context", "Application", "App", "Host", "Shell", "Project", "CurrentProject", "Workspace", "Manager"]
        try:
            objType = obj.GetType()
            for member_name in member_names:
                try:
                    prop = objType.GetProperty(member_name)
                    if prop is not None:
                        value = prop.GetValue(obj, None)
                        found = self._search_object_for_service(value, serviceType, depth + 1, visited)
                        if found is not None:
                            return found
                except:
                    pass
                try:
                    field = objType.GetField(member_name)
                    if field is not None:
                        value = field.GetValue(obj)
                        found = self._search_object_for_service(value, serviceType, depth + 1, visited)
                        if found is not None:
                            return found
                except:
                    pass
        except:
            pass

        return None

    def _try_construct_with_resolved_parameters(self, serviceType):
        try:
            ctors = serviceType.GetConstructors()
        except:
            return None

        ctor_list = [c for c in ctors]
        ctor_list.sort(key=lambda c: len(c.GetParameters()))
        for ctor in ctor_list:
            try:
                pars = ctor.GetParameters()
                args = []
                ok = True
                for p in pars:
                    built, value = self._try_resolve_runtime_object(p.ParameterType, "")
                    if not built:
                        ok = False
                        break
                    args.append(value)
                if ok:
                    return ctor.Invoke(Array[object](args))
            except:
                pass
        return None

    def _create_import_item_instance(self, itemTypeOrInterface, filePath):
        concreteType = self._resolve_concrete_import_item_type(itemTypeOrInterface)
        if concreteType is None:
            raise Exception("Could not resolve a concrete import item type for {0}".format(itemTypeOrInterface))

        try:
            item = Activator.CreateInstance(concreteType, Array[object]([filePath]))
            self._set_import_item_defaults(item)
            return item
        except:
            pass

        item = self._try_construct_with_resolved_parameters(concreteType)
        if item is None:
            item = Activator.CreateInstance(concreteType)

        self._set_import_item_path(item, filePath)
        self._set_import_item_defaults(item)
        return item

    def _resolve_concrete_import_item_type(self, itemTypeOrInterface):
        try:
            if itemTypeOrInterface.IsClass and (not itemTypeOrInterface.IsAbstract):
                return itemTypeOrInterface
        except:
            pass

        preferred_name_words = ["Import", "File", "Item"]
        candidates = []

        for asm in AppDomain.CurrentDomain.GetAssemblies():
            try:
                for t in asm.GetTypes():
                    try:
                        if (not t.IsClass) or t.IsAbstract:
                            continue
                        if itemTypeOrInterface.IsAssignableFrom(t):
                            candidates.append(t)
                    except:
                        pass
            except:
                pass

        if len(candidates) > 0:
            candidates.sort(key=lambda t: 0 if any(w in (t.Name or "") for w in preferred_name_words) else 1)
            return candidates[0]

        likely_names = [
            "ImportFilesServiceItem",
            "ImportFileServiceItem",
            "ImportServiceItem",
        ]
        for asm in AppDomain.CurrentDomain.GetAssemblies():
            try:
                for t in asm.GetTypes():
                    try:
                        if (not t.IsClass) or t.IsAbstract:
                            continue
                        if t.Name in likely_names:
                            return t
                    except:
                        pass
            except:
                pass

        return None

    def _try_project_object_import(self, jxlPath):
        errors = []
        try:
            target = self._currentProject
            if target is None:
                return False, None

            targetType = target.GetType()
            for method in targetType.GetMethods():
                try:
                    name = method.Name or ""
                    if "Import" not in name:
                        continue

                    pars = method.GetParameters()
                    can_build, args = self._build_method_arguments(pars, jxlPath)
                    if not can_build:
                        continue

                    method.Invoke(target, args)
                    return True, None
                except Exception as ex:
                    errors.append("currentProject." + self._format_method_signature(method) + ": " + self._unwrap_invoke_exception(ex))
        except Exception as ex:
            errors.append("currentProject import reflection: " + str(ex))

        if len(errors) > 0:
            return False, "; ".join(errors)
        return False, None

    def _set_import_item_path(self, item, filePath):
        t = item.GetType()

        for propName in ["FileName", "Filename", "Path", "FullPath", "SourceFileName"]:
            try:
                pInfo = t.GetProperty(propName)
                if pInfo is not None and pInfo.CanWrite:
                    pInfo.SetValue(item, filePath, None)
                    return
            except:
                pass

        for fieldName in ["FileName", "Filename", "Path", "FullPath", "SourceFileName"]:
            try:
                fInfo = t.GetField(fieldName)
                if fInfo is not None:
                    fInfo.SetValue(item, filePath)
                    return
            except:
                pass

        raise Exception("Could not set import file path on type {0}".format(t.FullName))

    def _invoke_service_method(self, importServiceType, method, args):
        target = None
        if not method.IsStatic:
            target = self._resolve_service_instance(importServiceType)
            if target is None:
                raise Exception("Could not resolve a live ImportService instance.")
        method.Invoke(target, args)

    def _set_import_item_defaults(self, item):
        try:
            t = item.GetType()
            for propName, value in [
                ("ShowReportAfterwards", True),
                ("ShowReportAfterImport", True),
                ("ZoomExtentsAfterwards", False),
                ("ZoomExtentsAfterImport", False),
            ]:
                pInfo = t.GetProperty(propName)
                if pInfo is not None and pInfo.CanWrite:
                    pInfo.SetValue(item, value, None)
        except:
            pass

    def _unwrap_invoke_exception(self, ex):
        try:
            inner = ex.InnerException
            if inner is not None:
                return str(inner)
        except:
            pass
        return str(ex)

    # ---------------------------
    # Close
    # ---------------------------

    def _close(self, sender, args):
        try:
            self.Close()
        except:
            pass
