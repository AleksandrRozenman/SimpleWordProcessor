/**
 * Programmer: Aleksandr Rozenman
 *
 * Simple word processor using Windows Forms.
 * Displays and saves .RTF files.
 * Allows user to customize the font, font size, font color.
 * Also allows user to use bullet points.
 * Provides the following common features:
 * cut/copy/paste, and select all, undo, redo.
 *
 * SimpleWordProcessor is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */

using System;
using System.Windows.Forms;

class SimpleWordProcessor : Form
{
	RichTextBox textbox;
	MainMenu mainmenu;
	FontDialog fDiag;
	ColorDialog cDiag;
	SaveFileDialog sDiag;
	OpenFileDialog oDiag;
	bool textChangedFlag;  // True means unsaved changes exist; false means file is up-to-date.
	string fileName;       // Currently active file; is null if working on a new file that has never been saved.

	public SimpleWordProcessor() : base()
	{
		textbox = new RichTextBox();
		textbox.Parent = this;
		textbox.Dock = DockStyle.Fill;
		textbox.AcceptsTab = true;
		
		textbox.TextChanged += delegate(object sender, EventArgs args) {
			textChangedFlag = true;
		};
		
		// Display save prompt if closing with unsaved data.
		this.FormClosing += delegate(object sender, FormClosingEventArgs args) {
			if(!confirmSave())
				args.Cancel = true;
		};
		
		this.Text = "SimpleWordProcessor";
		this.Width = 600;
		this.Height = 480;
		
		fDiag = new FontDialog();
		cDiag = new ColorDialog();
		sDiag = new SaveFileDialog();
		oDiag = new OpenFileDialog();
		
		textChangedFlag = false;
		
		buildMenu();
	}
	
	private void buildMenu()
	{
		// Create menu bar.
		mainmenu = new MainMenu();
		this.Menu = mainmenu;
		
		MenuItem miFile = new MenuItem();
		MenuItem miEdit = new MenuItem();
		MenuItem miStyle = new MenuItem();

		miFile.Text = "File";
		miEdit.Text = "Edit";
		miStyle.Text = "Style";

		mainmenu.MenuItems.Add(miFile);
		mainmenu.MenuItems.Add(miEdit);
		mainmenu.MenuItems.Add(miStyle);
		
		// Populate "File" menu.
		MenuItem fileNew = new MenuItem();
		MenuItem fileOpen = new MenuItem();
		MenuItem fileSave = new MenuItem();
		MenuItem fileSaveAs = new MenuItem();
		MenuItem fileExit = new MenuItem();
		
		fileNew.Text = "New";
		fileOpen.Text = "Open...";
		fileSave.Text = "Save";
		fileSaveAs.Text = "Save As...";
		fileExit.Text = "Exit";
		
		fileNew.Shortcut = Shortcut.CtrlN;
		fileOpen.Shortcut = Shortcut.CtrlO;
		fileSave.Shortcut = Shortcut.CtrlS;
		fileExit.Shortcut = Shortcut.AltF4;
		
		// For "New" and "Open" prompt user to save any unsaved work.
		fileNew.Click += delegate(object sender, EventArgs args) {
			if(confirmSave())
				newFile();
		};
		fileOpen.Click += delegate(object sender, EventArgs args) {
			if(confirmSave())
				open();
		};
		fileSave.Click += delegate(object sender, EventArgs args) {
			save();
		};
		fileSaveAs.Click += delegate(object sender, EventArgs args) {
			saveAs();
		};
		fileExit.Click += delegate(object sender, EventArgs args) {
			this.Close();
		};
		
		miFile.MenuItems.Add(fileNew);
		miFile.MenuItems.Add(fileOpen);
		miFile.MenuItems.Add(fileSave);
		miFile.MenuItems.Add(fileSaveAs);
		miFile.MenuItems.Add(new MenuItem("-"));  // Separator.
		miFile.MenuItems.Add(fileExit);
		
		// Populate "Edit" menu.
		MenuItem editCut = new MenuItem();
		MenuItem editCopy = new MenuItem();
		MenuItem editPaste = new MenuItem();
		MenuItem editUndo = new MenuItem();
		MenuItem editRedo = new MenuItem();
		MenuItem editSelectAll = new MenuItem();
		
		editCut.Text = "Cut";
		editCopy.Text = "Copy";
		editPaste.Text = "Paste";
		editUndo.Text = "Undo";
		editRedo.Text = "Redo";
		editSelectAll.Text = "Select All";
		
		editCut.Shortcut = Shortcut.CtrlX;
		editCopy.Shortcut = Shortcut.CtrlC;
		editPaste.Shortcut = Shortcut.CtrlV;
		editUndo.Shortcut = Shortcut.CtrlZ;
		editRedo.Shortcut = Shortcut.CtrlY;
		editSelectAll.Shortcut = Shortcut.CtrlA;
		
		miEdit.MenuItems.Add(editCut);
		miEdit.MenuItems.Add(editCopy);
		miEdit.MenuItems.Add(editPaste);
		miEdit.MenuItems.Add(editUndo);
		miEdit.MenuItems.Add(editRedo);
		miEdit.MenuItems.Add(editSelectAll);
		
		editCut.Click += delegate(object sender, EventArgs args) {
			if(textbox.SelectionLength > 0)
				textbox.Cut();
		};
		editCopy.Click += delegate(object sender, EventArgs args) {
			if(textbox.SelectionLength > 0)
				textbox.Copy();
		};
		editPaste.Click += delegate(object sender, EventArgs args) {
			// Including the following commented out line produces ThreadStateException
			//if(Clipboard.GetDataObject().GetDataPresent(DataFormats.Text) == true)
				textbox.Paste();
		};
		editUndo.Click += delegate(object sender, EventArgs args) {
			if(textbox.CanUndo)
				textbox.Undo();
		};
		editRedo.Click += delegate(object sender, EventArgs args) {
			if(textbox.CanRedo)
				textbox.Redo();
		};
		editSelectAll.Click += delegate(object sender, EventArgs args) {
			textbox.SelectionStart = 0;
			textbox.SelectionLength = textbox.TextLength;
		};
		
		// Populate "Style" menu.
		MenuItem styleFont = new MenuItem();
		MenuItem styleColor = new MenuItem();
		MenuItem styleBulletize = new MenuItem();
		
		styleFont.Text = "Font";
		styleColor.Text = "Color";
		styleBulletize.Text = "Bulletize";
		
		miStyle.MenuItems.Add(styleFont);
		miStyle.MenuItems.Add(styleColor);
		miStyle.MenuItems.Add(styleBulletize);
		
		styleFont.Click += delegate(object sender, EventArgs args) {
			fDiag.ShowDialog();
			textbox.SelectionFont = fDiag.Font;
		};
		styleColor.Click += delegate(object sender, EventArgs args) {
			cDiag.ShowDialog();
			textbox.SelectionColor = cDiag.Color;
		};
		styleBulletize.Click += delegate(object sender, EventArgs args) {
			textbox.SelectionBullet = !textbox.SelectionBullet;
		};
	}
	
	private void newFile()
	{
		textbox.Clear();
		textChangedFlag = false;
		fileName = null;
	}
	
	private void open()
	{
		oDiag.DefaultExt = "*.rtf";
		oDiag.Filter = "RTF Files|*.rtf";
		
		if(oDiag.ShowDialog() == DialogResult.OK && oDiag.FileName.Length > 0)
		{
			string fName = oDiag.FileName;
			if(openFile(fName))
			{
				textChangedFlag = false;
				fileName = fName;
			}
		}
	}
	
	// Returns true if file was opened successfully, false otherwise.
	// If file was not opened successfully, display error message.
	private bool openFile(string fName)
	{
		try
		{
			textbox.LoadFile(fName);
			return true;
		}
		catch(Exception ex)
		{
			// Open failed. Prompt user of failure, and ask to either retry opening or cancel attempt.
			string message = "The file could not be opened. It may be in use by another program."
			                 + "\nTry closing any program that may be using the file and try again.";
			string caption = "I Am Error";
			MessageBoxButtons mbButtons = MessageBoxButtons.RetryCancel;
			DialogResult result;
			
			// Read user's choice.
			result = MessageBox.Show(message, caption, mbButtons);
			// I want to try opening again. Repeat this. Return true if opening eventually succeeded.
			if(result == DialogResult.Retry)
				return openFile(fName);
			
			// User eventually gave up opening.
			return false;
		}
	}
	
	private void save()
	{
		// If the file to save to does not exist,
		// (e.g., was deleted while editing, or a brand new file was created)
		// prompt for a file name, else go ahead and save without further user input.
		if(!System.IO.File.Exists(fileName))
			saveAs();
		else if(textChangedFlag)	// Don't waste time saving if there's nothing to save.
		{
			// Only set textChangedFlag to false if file was successfully saved.
			// (Since if the file wasn't saved successfully, then it's still outdated.)
			if(saveFile(fileName))
				textChangedFlag = false;
		}
	}
	
	// Same as save, only prompts user to enter a target filename.
	private void saveAs()
	{
		sDiag.DefaultExt = "*.rtf";
		sDiag.Filter = "RTF Files|*.rtf";

		if(sDiag.ShowDialog() == DialogResult.OK && sDiag.FileName.Length > 0)
		{
			string fName = sDiag.FileName;
			if(saveFile(fName))
			{
				textChangedFlag = false;
				fileName = fName;
			}
		}
	}
	
	// Returns true if file was saved successfully, false otherwise.
	// If file was not saved successfully, display error message.
	private bool saveFile(string fName)
	{
		try
		{
			textbox.SaveFile(fName);
			return true;
		}
		catch(Exception ex)
		{
			// Save failed. Prompt user of failure, and ask to either retry saving or cancel attempt.
			string message = "The file could not be written. It may be in use by another program."
			                 + "\nTry closing any program that may be using the file and try again.";
			string caption = "I Am Error";
			MessageBoxButtons mbButtons = MessageBoxButtons.RetryCancel;
			DialogResult result;
			
			// Read user's choice.
			result = MessageBox.Show(message, caption, mbButtons);
			// I want to try saving again. Repeat this. Return true if saving eventually succeeded.
			if(result == DialogResult.Retry)
				return saveFile(fName);
			
			// User eventually gave up saving.
			return false;
		}
	}
	
	// Returns false if cancel clicked; true otherwise (even if not saved).
	private bool confirmSave()
	{
		string message;
		
		// Nothing to save if nothing was changed.
		// Therefore, can go ahead and do whatever it was that was going to be done.
		if(!textChangedFlag)
			return true;
		
		// Prompt user to save changes before closing the current document or program.
		if(fileName == null)
			message = "The file has been modified.\nDo you want to save your changes?";
		else
			message = "The file \"" + fileName + "\" has been modified.\nDo you want to save your changes?";
		string caption = "Save?";
		MessageBoxButtons mbButtons = MessageBoxButtons.YesNoCancel;
		DialogResult result;
		
		result = MessageBox.Show(message, caption, mbButtons);
		
		// I do not want to do what I said I was doing, take me back.
		if(result == DialogResult.Cancel)
			return false;

		// I want to save before doing what I said I wanted to do.
		if(result == DialogResult.Yes)
			save();
		// No else statement. Else would have been empty anyway.
		// It meant that I do not want to save before doing whatever I said to do;
		// discard all changes I made.
		
		// File either saved or I indicated that I want to discard changes.
		// In either case, continue doing whatever it is I wanted to do.
		return true;
	}
		
	[STAThread]
	public static void Main()
	{
		SimpleWordProcessor form = new SimpleWordProcessor();
		Application.EnableVisualStyles();
		Application.Run(form);
	}
}
