using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualBasic;
using Microsoft.VisualStudio.CommandBars;

namespace Emmet4VsAddin
{
    /// <summary>
    /// </summary>
    [DataContract]
    public class Action
    {
        /// <summary>
        /// </summary>
        [DataMember]
        public string type { get; set; }
        /// <summary>
        /// </summary>
        [DataMember]
        public string name { get; set; }
        /// <summary>
        /// </summary>
        [DataMember]
        public string label { get; set; }
        /// <summary>
        /// </summary>
        [DataMember]
        public Action[] items { get; set; }
    }

    /// <summary>
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Context
    {
        private DTE2 _dte = null;
        private AddIn _addIn;
        private ScriptEngine _engine = null;
        private string _root = null;
        
        /// <summary>
        /// 
        /// </summary>
        public string Root
        {
            get
            {
                if (_root == null)
                {
                    Uri uri = new Uri(Assembly.GetExecutingAssembly().EscapedCodeBase);
                    string temp = System.IO.Path.GetDirectoryName(Uri.UnescapeDataString(uri.AbsolutePath));
                    _root = System.IO.Directory.GetParent(temp).FullName;
                }
                return _root;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Path
        {
            get
            {
                return _dte.ActiveDocument.FullName;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Text
        {
            get
            {
                TextDocument doc = _dte.ActiveDocument.Object("TextDocument") as TextDocument;
                return doc.StartPoint.CreateEditPoint().GetText(doc.EndPoint);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Line
        {
            get
            {
                TextSelection selection = _dte.ActiveDocument.Selection as TextSelection;
                VirtualPoint active = selection.ActivePoint;
                EditPoint point0 = active.CreateEditPoint();
                EditPoint point1 = active.CreateEditPoint();
                point0.StartOfLine();
                point1.EndOfLine();
                return point0.GetText(point1);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Anchor
        {
            get
            {
                TextSelection selection = _dte.ActiveDocument.Selection as TextSelection;
                return locate(selection.AnchorPoint.AbsoluteCharOffset);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Active
        {
            get
            {
                TextSelection selection = _dte.ActiveDocument.Selection as TextSelection;
                return locate(selection.ActivePoint.AbsoluteCharOffset);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public int LineBegOffset
        {
            get
            {
                TextSelection selection = _dte.ActiveDocument.Selection as TextSelection;
                VirtualPoint active = selection.ActivePoint;
                EditPoint beg = active.CreateEditPoint();
                beg.StartOfLine();
                return locate(beg.AbsoluteCharOffset);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public int LineEndOffset
        {
            get
            {
                TextSelection selection = _dte.ActiveDocument.Selection as TextSelection;
                VirtualPoint active = selection.ActivePoint;
                EditPoint end = active.CreateEditPoint();
                end.EndOfLine();
                return locate(end.AbsoluteCharOffset);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Syntax
        {
            get
            {
                return "html";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Profile
        {
            get
            {
                return "html";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Selection
        {
            get
            {
                TextSelection selection = _dte.ActiveDocument.Selection as TextSelection;
                return selection.Text;
            }
            set
            {
                TextSelection selection = _dte.ActiveDocument.Selection as TextSelection;
                selection.Text = value;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dte"></param>
        /// <param name="addIn"></param>
        /// <param name="engine"></param>
        /// <returns></returns>
        public Context(DTE2 dte, AddIn addIn, ScriptEngine engine)
        {
            _dte = dte;
            _addIn = addIn;
            _engine = engine;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="beg"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public bool select(int beg, int end)
        {
            TextSelection selection = _dte.ActiveDocument.Selection as TextSelection;
            selection.MoveToAbsoluteOffset(offset(beg), false);
            if (beg != end)
            {
                selection.MoveToAbsoluteOffset(offset(end), true);
                return true;
            }
            return false;
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="beg"></param>
        /// <param name="end"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public bool replace(int beg, int end, string value)
        {
            TextSelection selection = _dte.ActiveDocument.Selection as TextSelection;
            VirtualPoint active = selection.ActivePoint;
            EditPoint point0 = active.CreateEditPoint();
            EditPoint point1 = active.CreateEditPoint();
            point0.MoveToAbsoluteOffset(offset(beg));
            point1.MoveToAbsoluteOffset(offset(end));
            point0.ReplaceText(point1, value, 1);
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public bool require(string path)
        {
            try
            {
                _engine.Exec(path);
            }
            catch
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="json"></param>
        /// <returns></returns>
        public bool startup(string json)
        {
            object[] contextGUIDS = new object[] { };
            Commands2 commands = _dte.Commands as Commands2;
            CommandBar menuBar = (_dte.CommandBars as CommandBars)["MenuBar"];

            try
            {
                int index = 9;
                foreach (CommandBarControl control in menuBar.Controls)
                {
                    if (control.Type == MsoControlType.msoControlPopup)
                    {
                        CommandBarPopup temp = control as CommandBarPopup;
                        if (temp.CommandBar.Name == "Tools")
                        {
                            index = temp.Index + 1;
                            break;
                        }
                    }
                }
                CommandBarPopup popup = menuBar.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, Type.Missing, index, true) as CommandBarPopup;
                popup.CommandBar.Name = "Emmet";
                popup.Caption = "E&mmet";
                CommandBar commandBar = popup.CommandBar;
                Command command = null;
                CommandBarButton button = null;

                DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Action[]));
                MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(json));
                Action[] actions = serializer.ReadObject(stream) as Action[];

                foreach (var action in actions)
                {
                    if (action.type == "action")
                    {
                        command = commands.AddNamedCommand2(_addIn, action.name, action.label, action.label, false, Type.Missing, ref contextGUIDS, 1);
                        button = command.AddControl(commandBar, commandBar.Controls.Count + 1) as CommandBarButton;
                    }
                    else
                    {
                        CommandBar control = commands.AddCommandBar(action.name, vsCommandBarType.vsCommandBarTypeMenu, commandBar, commandBar.Controls.Count + 1) as CommandBar;
                        if (action.items != null)
                        {
                            foreach (var item in action.items)
                            {
                                command = commands.AddNamedCommand2(_addIn, item.name, item.label, item.label, false, Type.Missing, ref contextGUIDS, 1);
                                button = command.AddControl(control, control.Controls.Count + 1) as CommandBarButton;
                            }
                        }
                    }
                }
                popup.Visible = true;
            }
            catch
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="title"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public string prompt(string title, string value)
        {
            return Interaction.InputBox(title, title, null);
        }

        private int locate(int off)
        {
            TextSelection selection = _dte.ActiveDocument.Selection as TextSelection;
            VirtualPoint active = selection.ActivePoint;
            EditPoint point0 = active.CreateEditPoint();
            EditPoint point1 = active.CreateEditPoint();
            point0.StartOfDocument();
            point1.MoveToAbsoluteOffset(off);
            string text = point0.GetText(point1);
            return (off - 1) + length(text, "\r\n", text.Length - 1);
        }

        private int offset(int pos)
        {
            return (pos + 1) + length(Text, "\r\n", pos);
        }

        private int length(string text, string omit, int last)
        {
            int num = 0;
            int idx = 0;
            int len = omit.Length;
            while (true)
            {
                int tmp = text.IndexOf(omit, idx, StringComparison.Ordinal);
                if (tmp == -1 || tmp >= last)
                {
                    return num;
                }
                num++;
                idx = tmp + len;
            }
        }
    }
}
