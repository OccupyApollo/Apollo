using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace SystemHotKeysLib
{
    /// Contains windows APIs used in this library.
    class Api
    {
        /// Registers a system hot key.
        /// <param name="hWnd">Handle of the window.</param>
        /// <param name="id">Hot key identifier.</param>
        /// <param name="fsModifiers">Key modifiers.</param>
        /// <param name="vk">Virtual key code.</param>
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool RegisterHotKey(
            IntPtr hWnd,
            int id,
            uint fsModifiers,
            Keys vk
        );

        /// Unregisters a system hot key.
        /// <param name="hWnd">Handle of the window.</param>
        /// <param name="id">Hot key identifier.</param>
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(
            IntPtr hWnd,
            int id
        );
        /// The windows message that indicates a hot key was pressed.
        public const int WM_HOTKEY = 0x0312;
    }

    /// Represents a system hot key.
    public class HotKey
    {
        /// Initializes a new instance of the <see cref="HotKey"/> by specific name, modifiers, and key.
        /// <param name="name">A <see cref="string"/> to identify the system hot key.</param>
        /// <param name="modifiers">Modify keys (Shift, Ctrl, Alt, and/or Windows Logo key).</param>
        /// <param name="key">Key code.</param>
        public HotKey(string name, KeyModifiers modifiers, Keys key)
        {
            if (name == null)
            {
                throw new ArgumentNullException("name");
            }

            if (key == Keys.None)
            {
                throw new ArgumentException("key cannot be Keys.None. Must specify a key for the hot key.", "key");
            }

            _name = name;
            _modifiers = modifiers;
            _key = key;
        }

        private int _id = 0;
        /// Gets or sets the ID of the hot key.
        internal int Id
        {
            get { return _id; }
            set { _id = value; }
        }

        private string _name = null;
        /// A <see cref="string"/> used to identify the hot key.
        public string Name
        {
            get { return _name; }
        }

        private KeyModifiers _modifiers = KeyModifiers.None;
        /// Modifier keys of this hot key.
        public KeyModifiers Modifiers
        {
            get { return _modifiers; }
        }

        private Keys _key = Keys.None;
        /// Key of this hot key.
        public Keys Key
        {
            get { return _key; }
        }
    }

    /// Modifier keys.
    [Flags]
    public enum KeyModifiers
    {
        /// No modifier key was pressed.
        None = 0,
        /// Alt key was pressed.
        Alt = 1,
        /// Ctrl key was pressed.
        Control = 2,
        /// Shift key was pressed.
        Shift = 4,
        /// Windows Logo key was pressed.
        Windows = 8
    }

    /// Processes the system hot key message.
    internal class SystemHotKeyMessageFilter : IMessageFilter
    {
        /// Initialize a new instance of the <see cref="SystemHotKeyMessageFilter"/> with
        /// the specific <see cref="SystemHotKeyListener"/>
        /// <param name="listner">An instance of the <see cref="SystemHotKeyListener"/>.</param>
        public SystemHotKeyMessageFilter(SystemHotKeyListener listner)
        {
            if (listner == null)
            {
                throw new ArgumentNullException("listner");
            }

            _listner = listner;
        }

        private SystemHotKeyListener _listner = null;

        #region IMessageFilter Members

        public bool PreFilterMessage(ref Message m)
        {
            switch (m.Msg)
            {
                case Api.WM_HOTKEY:
                    // Get the system hot key message.
                    ProcessHotkey(m);
                    break;
            }

            return false;
        }

        #endregion

        private void ProcessHotkey(Message m)
        {
            // Get internal id of the system hot key.
            int sid = m.WParam.ToInt32();

            // Check if it is registered in this program.
            if (_listner._hotKeys.ContainsKey(sid))
            {
                // If yes, invoke the event.
                HotKey hotKey = _listner._hotKeys[sid];
                _listner.InvokeSystemHotKeyPressed(hotKey);
            }
        }

    }

    /// <summary>
    /// Registers, listens, and unregisters system hot keys.
    /// </summary>
    public class SystemHotKeyListener : IDisposable
    {
        /// Occurs when the registered system hot key was pressed.
        public event SystemHotKeyPressedEventHandler SystemHotKeyPressed;

        /// Initializes a new instance of the <see cref="SystemHotKeyListener"/>.
        public SystemHotKeyListener()
        {
            _listeningForm = new Form();
            _listeningForm.Visible = false;

            _messageFilter = new SystemHotKeyMessageFilter(this);
            Application.AddMessageFilter(_messageFilter);
        }

        private Form _listeningForm = null;

        internal Dictionary<int, HotKey> _hotKeys = new Dictionary<int, HotKey>();
        public ICollection<HotKey> HotKeys
        {
            get { return _hotKeys.Values as ICollection<HotKey>; }
        }

        private SystemHotKeyMessageFilter _messageFilter = null;

        private int nextId = 100;

        /// Registers a new system hot key.
        /// <param name="hotKey">The <see cref="HotKey"/> that need to register.</param>
        /// <exception cref="InvalidOperationException">
        /// The same hot key or name already exists.
        /// </exception>
        public void Register(HotKey hotKey)
        {
            // Check for duplicate hot key name.
            foreach (HotKey existingHotKey in _hotKeys.Values)
            {
                if (existingHotKey.Name == hotKey.Name)
                {
                    throw new InvalidOperationException("The hot key with name of '" + hotKey.Name + "' has been already registered.");
                }

                if (existingHotKey.Modifiers == hotKey.Modifiers &&
                    existingHotKey.Key == hotKey.Key)
                {
                    throw new InvalidOperationException("The same hot key has been already registered.");
                }
            }

            // Register the hot key.
            hotKey.Id = nextId;
            _hotKeys.Add(nextId, hotKey);
            Api.RegisterHotKey(_listeningForm.Handle, nextId, (uint)(hotKey.Modifiers), hotKey.Key);

            // Increase the internal hot key id.
            nextId++;
        }

        /// <summary>
        /// Unregisters the specific hot key.
        /// </summary>
        /// <param name="hotKey">The <see cref="HotKey"/> that needs to be unregistered.</param>
        /// <returns>A <see cref="bool"/> value that indicates whether the operation is successed.</returns>
        public bool Unregister(HotKey hotKey)
        {
            if (_hotKeys.ContainsKey(hotKey.Id))
            {
                Api.UnregisterHotKey(_listeningForm.Handle, hotKey.Id);
                _hotKeys.Remove(hotKey.Id);
                return true;
            }
            else
            {
                return false;
            }
        }

        /// Unregisters all system hot keys registered by this program.
        public void UnregisterAll()
        {
            int[] idList = new int[_hotKeys.Count];
            _hotKeys.Keys.CopyTo(idList, 0);

            for (int i = 0; i < idList.Length; i++)
            {
                int id = idList[i];

                Api.UnregisterHotKey(_listeningForm.Handle, id);
                _hotKeys.Remove(id);
            }
        }

        protected HotKey _lastPressedHotKey = null;
        protected DateTime _lastPressedTime = DateTime.MinValue;

        internal void InvokeSystemHotKeyPressed(HotKey hotKey)
        {
            SystemHotKeyPressedEventArgs e = new SystemHotKeyPressedEventArgs(hotKey);

            // Single click
            _lastPressedHotKey = hotKey;
            _lastPressedTime = DateTime.Now;

            OnSystemHotKeyPressed(e);
        }

        protected virtual void OnSystemHotKeyPressed(SystemHotKeyPressedEventArgs e)
        {
            if (SystemHotKeyPressed != null)
            {
                SystemHotKeyPressedEventHandler invoker = SystemHotKeyPressed;
                invoker(this, e);
            }
        }

        #region IDisposable Members

        /// Releases all the resources used by the <see cref="SystemHotKeyListener"/>.
        public void Dispose()
        {
            UnregisterAll();
            _listeningForm.Dispose();
        }

        #endregion
    }

    /// Provides data for the SystemHotKeyListener.SystemHotKeyPressed event.
    public class SystemHotKeyPressedEventArgs : EventArgs
    {
        /// Initializes a new instance of the <see cref="SystemHotKeyPressedEventArgs"/> class.
        /// <param name="hotKey">A <see cref="HotKey"/> instance that specifies which hot key was pressed.</param>
        public SystemHotKeyPressedEventArgs(HotKey hotKey)
        {
            _hotKey = hotKey;
        }

        private HotKey _hotKey = null;
        /// Gets the hot key that was just pressed.
        public HotKey PressedHotKey
        {
            get { return _hotKey; }
        }

        /// Gets the name of the hot key that was just pressed.
        public string HotKeyName
        {
            get { return _hotKey.Name; }
        }
    }

    public delegate void SystemHotKeyPressedEventHandler(
        object sender,
        SystemHotKeyPressedEventArgs e);
}
