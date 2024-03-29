using System;
using System.Windows.Forms;
using System.ComponentModel;
using System.Xml.Serialization;
using System.Runtime.InteropServices;

// Based on https://bloggablea.wordpress.com/2007/05/01/global-hotkeys-with-net/
namespace MonitorSwitcherGUI
{
    public class HotkeyCtrl : IMessageFilter
	{
		#region Interop

		[DllImport("user32.dll", SetLastError = true)]
		private static extern int RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, Keys vk);

		[DllImport("user32.dll", SetLastError=true)]
		private static extern int UnregisterHotKey(IntPtr hWnd, int id);

		private const uint WM_HOTKEY = 0x312;

		private const uint MOD_ALT = 0x1;
		private const uint MOD_CONTROL = 0x2;
		private const uint MOD_SHIFT = 0x4;
		private const uint MOD_WIN = 0x8;

		private const uint ERROR_HOTKEY_ALREADY_REGISTERED = 1409;

		#endregion

		private static int currentID;
		private const int maximumID = 0xBFFF;
		
		private Keys keyCode;
        private bool shift;
        private bool control;
        private bool alt;
		private bool windows;

		[XmlIgnore]
		private int id;
		[XmlIgnore]
		private bool registered = false;
		[XmlIgnore]
		private Control windowControl;

		public event HandledEventHandler Pressed;

		public HotkeyCtrl() : this(Keys.None, false, false, false, false)
		{
			// No work done here!
		}
		
		public HotkeyCtrl(Keys keyCode, bool shift, bool control, bool alt, bool windows)
		{
			// Assign properties
			this.KeyCode = keyCode;
			this.Shift = shift;
			this.Control = control;
			this.Alt = alt;
			this.Windows = windows;

			// Register us as a message filter
			Application.AddMessageFilter(this);
		}

        ~HotkeyCtrl()
		{
			// Unregister the hotkey if necessary
			if (this.Registered)
			{ this.Unregister(); }
		}

        public HotkeyCtrl Clone()
		{
			// Clone the whole object
            return new HotkeyCtrl(this.keyCode, this.shift, this.control, this.alt, this.windows);
		}

		public bool GetCanRegister(Control windowControl)
		{
			// Handle any exceptions: they mean "no, you can't register" :)
			try
			{
				// Attempt to register
				if (!this.Register(windowControl))
				{ return false; }

				// Unregister and say we managed it
				this.Unregister();
				return true;
			}
			catch (Win32Exception)
			{ return false; }
			catch (NotSupportedException)
			{ return false; }
		}

		public bool Register(Control windowControl)
        {
            // Check that we have not registered
			if (this.registered)
			{ throw new NotSupportedException("You cannot register a hotkey that is already registered"); }
        
			// We can't register an empty hotkey
			if (this.Empty)
			{ throw new NotSupportedException("You cannot register an empty hotkey"); }

			// Get an ID for the hotkey and increase current ID
            this.id = HotkeyCtrl.currentID;
            HotkeyCtrl.currentID = HotkeyCtrl.currentID + 1 % HotkeyCtrl.maximumID;

			// Translate modifier keys into unmanaged version
            uint modifiers = (this.Alt ? HotkeyCtrl.MOD_ALT : 0) | (this.Control ? HotkeyCtrl.MOD_CONTROL : 0) |
                            (this.Shift ? HotkeyCtrl.MOD_SHIFT : 0) | (this.Windows ? HotkeyCtrl.MOD_WIN : 0);

			// Register the hotkey
            if (HotkeyCtrl.RegisterHotKey(windowControl.Handle, this.id, modifiers, keyCode) == 0)
			{ 
				// Is the error that the hotkey is registered?
				if (Marshal.GetLastWin32Error() == ERROR_HOTKEY_ALREADY_REGISTERED)
				{ return false; }
				else
				{ throw new Win32Exception(); } 
			}

			// Save the control reference and register state
			this.registered = true;
			this.windowControl = windowControl;

			// We successfully registered
			return true;
		}

		public void Unregister()
		{
			// Check that we have registered
			if (!this.registered)
			{ throw new NotSupportedException("You cannot unregister a hotkey that is not registered"); }
        
			// It's possible that the control itself has died: in that case, no need to unregister!
			if (!this.windowControl.IsDisposed)
			{
				// Clean up after ourselves
                if (HotkeyCtrl.UnregisterHotKey(this.windowControl.Handle, this.id) == 0)
				{ 
                    //throw new Win32Exception(); 
                }
			}

			// Clear the control reference and register state
			this.registered = false;
			this.windowControl = null;
		}

		private void Reregister()
		{
			// Only do something if the key is already registered
			if (!this.registered)
			{ return; }

			// Save control reference
			Control windowControl = this.windowControl;

			// Unregister and then reregister again
			this.Unregister();
			this.Register(windowControl);
		}

		public bool PreFilterMessage(ref Message message)
		{
			// Only process WM_HOTKEY messages
            if (message.Msg != HotkeyCtrl.WM_HOTKEY)
			{ return false; }

			// Check that the ID is our key and we are registerd
			if (this.registered && (message.WParam.ToInt32() == this.id))
			{
				// Fire the event and pass on the event if our handlers didn't handle it
				return this.OnPressed();
			}
			else
			{ return false; }
		}

		private bool OnPressed()
		{
			// Fire the event if we can
			HandledEventArgs handledEventArgs = new HandledEventArgs(false);
			if (this.Pressed != null)
			{ this.Pressed(this, handledEventArgs); }

			// Return whether we handled the event or not
			return handledEventArgs.Handled;
		}

        public override string ToString()
        {
			// We can be empty
			if (this.Empty)
			{ return "(none)"; }

			// Build key name
			string keyName = Enum.GetName(typeof(Keys), this.keyCode);;
			switch (this.keyCode)
			{
				case Keys.D0:
				case Keys.D1:
				case Keys.D2:
				case Keys.D3:
				case Keys.D4:
				case Keys.D5:
				case Keys.D6:
				case Keys.D7:
				case Keys.D8:
				case Keys.D9:
					// Strip the first character
					keyName = keyName.Substring(1);
					break;
				default:
					// Leave everything alone
					break;
			}

            // Build modifiers
            string modifiers = "";
            if (this.shift)
            { modifiers += "Shift+"; }
            if (this.control)
            { modifiers += "Control+"; }
            if (this.alt)
            { modifiers += "Alt+"; }
			if (this.windows)
			{ modifiers += "Windows+"; }

			// Return result
            return modifiers + keyName;
        }

		public bool Empty => this.keyCode == Keys.None;

        public bool Registered => this.registered;

        public Keys KeyCode
        {
            get => this.keyCode;
            set
			{
				// Save and reregister
				this.keyCode = value;
				this.Reregister();
			}
        }

        public bool Shift
        {
            get => this.shift;
            set 
			{
				// Save and reregister
				this.shift = value;
				this.Reregister();
			}
        }

        public bool Control
        {
            get => this.control;
            set
			{ 
				// Save and reregister
				this.control = value;
				this.Reregister();
			}
        }

        public bool Alt
        {
            get => this.alt;
            set
			{ 
				// Save and reregister
				this.alt = value;
				this.Reregister();
			}
        }

		public bool Windows
		{
			get => this.windows;
			set 
			{
				// Save and reregister
				this.windows = value;
				this.Reregister();
			}
		}
    }
}
