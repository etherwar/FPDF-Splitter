import os
import sys
import wx
import PyPDF2

PdfFileWriter = PyPDF2.PdfFileWriter
PdfFileReader = PyPDF2.PdfFileReader

__version__ = '1.0'

####################################################################################################
######################################### MainWindow Class #########################################


class MainWindow(wx.Frame):

    PAD = 5

    def __init__(self, parent, title):
        self.src_dir: str = ''
        self.src_file: str = ''
        self.out_dir: str = ''
        self.prefix_mod = False

        ####################################################################################################
        ##################################### WINDOW AND PANEL SET UP ######################################

        # A "-1" in the size parameter instructs wxWidgets to use the default size.
        # In this case, we will provide the default size and set the width to 575.
        wx.Frame.__init__(self, parent, title=title) # (575, -1)
        # Set Sizing #######################################
        self.SetMinSize(wx.Size(335, 434))
        self.SetSize(wx.Size(580, 450))
        self.SetMaxSize(wx.Size(1920, 550))

        # STATUS BAR #######################################
        self.statusbar = self.CreateStatusBar() # A Statusbar in the bottom of the window
        self._statusbar()

        # FILE MENU ########################################
        # Setting up the File menu.
        filemenu = wx.Menu()
        menu_select = filemenu.Append(wx.ID_OPEN, "&Select PDF", " Open a file to process.")
        menu_sep = filemenu.Append(wx.MenuItem())
        menu_exit = filemenu.Append(wx.ID_EXIT, "E&xit", " Terminate the program.")

        helpmenu = wx.Menu()
        menu_about = helpmenu.Append(wx.ID_HELP, "&About", " About this program.")

        # Creating the menu bar ############################
        menu_bar = wx.MenuBar()
        menu_bar.Append(filemenu, "&File")  # Adding the "filemenu" to the MenuBar
        menu_bar.Append(helpmenu, "&Help")
        self.SetMenuBar(menu_bar)  # Adding the MenuBar to the Frame content.

        # MAIN FORM ########################################
        # Create the Form elements #########################
        # Create a panel from the frame
        self.pnl = wx.Panel(self, size=(580, -1))

        ####################################################
        # Initiate vertical column sizer for image sizer and flexgrid sizer
        self.sizer = wx.BoxSizer(orient=wx.VERTICAL)

        ####################################################
        # Create the 1x1 Image Panel #######################
        # HEADER ###########################################
        # Fix for for either running the program as a frozen (cx_Freeze) program or unfrozen
        if getattr(sys, 'frozen', False):
            # frozen
            dir_ = os.path.dirname(sys.executable)
        else:
            # unfrozen
            dir_ = os.path.dirname(os.path.realpath(__file__))

        bmp = wx.Bitmap(os.path.join(dir_, 'force.png'))
        head = wx.StaticBitmap(self.pnl, bitmap=bmp, pos=(0, 0), size=(550, 248))
        self.sizer.Add(head, 1, wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, self.PAD)

        # ROW 1 ############################################
        lbl_folder = wx.StaticText(self.pnl, label="Select PDF to Split: ")
        txt_filename_preview_style: int = wx.TE_LEFT | wx.TE_READONLY | wx.TE_NO_VSCROLL
        self.txt_filename_preview = wx.TextCtrl(self.pnl, style=txt_filename_preview_style)
        self.btn_openfile = wx.Button(self.pnl, id=wx.ID_ADD, label="Select PDF")

        # DRAG AND DROP FUNCTIONALITY FOR INITIAL FILE #####
        # TODO: Finish Drag and Drop
        self.txt_filename_preview.DragAcceptFiles(True)
        # drop_event = wx.DropFilesEvent

        # ROW 2 ############################################
        lbl_output_pre = wx.StaticText(self.pnl, label="Out File Prefix:")
        txt_file_prefix_style: int = wx.TE_LEFT | wx.TE_NO_VSCROLL
        self.txt_file_prefix = wx.TextCtrl(self.pnl, style=txt_file_prefix_style)
        lbl_output_suf = wx.StaticText(self.pnl, label="Suffix:")
        txt_file_suffix_style: int = wx.TE_LEFT | wx.TE_NO_VSCROLL
        self.txt_file_suffix = wx.TextCtrl(self.pnl, style=txt_file_suffix_style)
        self.btn_setoutdir = wx.Button(self.pnl, id=wx.ID_OPEN, label="Select Dir...")

        # ROW 3 ############################################
        lbl_outputprev = wx.StaticText(self.pnl, label="Sample Final Output:")
        txt_outputprev_style: int = wx.TE_LEFT | wx.TE_READONLY | wx.TE_NO_VSCROLL
        self.txt_outputprev = wx.TextCtrl(self.pnl, style=txt_outputprev_style)
        self.btn_process = wx.Button(self.pnl, id=wx.ID_APPLY)

        ####################################################
        # Binding ##########################################

        # Menu Event
        self.Bind(wx.EVT_MENU, self.on_open, menu_select)
        self.Bind(wx.EVT_MENU, self.on_exit, menu_exit)
        self.Bind(wx.EVT_MENU, self.on_about, menu_about)

        # Button Event
        self.btn_openfile.Bind(wx.EVT_BUTTON, self.on_open)
        self.btn_setoutdir.Bind(wx.EVT_BUTTON, self.on_select)
        self.btn_process.Bind(wx.EVT_BUTTON, self.on_process)

        ####################################################
        # AddMany() matrix simplfier for 3x3 flexgrid ######
        ####################################################
        # Objects ##########################################
        # Row 1  ###########################################
        topleft = (lbl_folder,)
        topmid = (self.txt_filename_preview,)
        topright = (self.btn_openfile,)

        # Row 2  ###########################################
        midleft = (lbl_output_pre,)

        # Create a row object with 3 columns (txt_pre, label_suffix, txt_suffix
        row = wx.BoxSizer()
        row.Add(self.txt_file_prefix, 1, wx.EXPAND)
        row.Add(lbl_output_suf, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_RIGHT | wx.LEFT, self.PAD)
        row.Add(self.txt_file_suffix, 1, wx.EXPAND | wx.LEFT, 2)

        mid = (row,)  # (self.txt_file_prefix,)
        midright = (self.btn_setoutdir,)

        # Row 3  ###########################################
        botleft = (lbl_outputprev,)
        botmid = (self.txt_outputprev,)
        botright = (self.btn_process,)

        # Styles ###########################################
        # (proportion=0, flag=0, border=0, userData=None) ##
        # Row 1  ###########################################
        style_topleft = (0, wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_RIGHT | wx.LEFT | wx.TOP, self.PAD)
        style_topmid = (1, wx.EXPAND | wx.TOP, self.PAD)
        style_topright = (0, wx.TOP | wx.RIGHT, self.PAD)

        # Row 2  ###########################################
        style_midleft = (0, wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_RIGHT | wx.LEFT, self.PAD)
        style_mid = (1, wx.EXPAND)
        style_midright = (0, wx.RIGHT, self.PAD)

        # Row 3  ###########################################
        style_botleft = (0, wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_RIGHT | wx.LEFT | wx.BOTTOM, self.PAD)
        style_botmid = (1, wx.EXPAND | wx.BOTTOM, self.PAD)
        style_botright = (0, wx.BOTTOM | wx.RIGHT, self.PAD)

        ####################################################
        # FlexGridSizer (simplified) #######################
        self.sizer2 = wx.FlexGridSizer(3, 3, 10, 2)       # (rows, cols, row_gap, col_gap)
        self.sizer2.AddMany((topleft + style_topleft,     # POS[0, 0] [x, y] cartesian position of flex grid elements
                            topmid + style_topmid,       # POS[0, 1] row 1 col 2
                            topright + style_topright,   # POS[0, 2] row 1 col 3 ... etc.
                            midleft + style_midleft,     # POS[1, 0] left to right, then up to down
                            mid + style_mid,             # POS[1, 1]
                            midright + style_midright,   # POS[1, 2]
                            botleft + style_botleft,     # POS[2, 0]
                            botmid + style_botmid,       # POS[2, 1]
                            botright + style_botright))  # POS[2, 2]

        self.sizer2.AddGrowableCol(1, proportion=1)

        self.sizer.Add(self.sizer2, 1, wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, self.PAD)

        # Layout sizers
        self.pnl.SetSizer(self.sizer)
        self.Fit()
        # self.pnl.SetAutoLayout(True)        # tells your window to use the sizer to position and size your components
        self.Show()

    ####################################################################################################
    ################################### PROPERTY GETTERS AND SETTERS ###################################

    @property
    def status(self):
        return self.GetStatusBar().GetStatusText()

    @status.setter
    def status(self, val: str):
        if issubclass(type(val), str):
            self.SetStatusText(val)
        else:
            raise ValueError("[-] Statusbar must be of subclass str.")

    ####################################################################################################
    ########################################## PRIVATE FUNCS ###########################################

    @staticmethod
    def _fnslug(filename: str, keep_chars: tuple = (' ', '.', '_')):
        return ''.join(c for c in filename if c.isalnum() or c in keep_chars)

    def _dialog(self, msg=None, title='Error', styles=None, pos=wx.DefaultPosition):
        """Dialog method that allows easy creation of 'popup' dialogs and
        defaults to a generic error message"""

        if msg is None:
            msg = "An unspecified error has occurred, fire the lazy programmer. Oh wait, this is free " \
                  "(just hurl insults instead)."

        if styles is None:
            styles = wx.OK | wx.CENTRE | wx.STAY_ON_TOP

        d = wx.MessageDialog(self, msg, title, style=styles, pos=pos)

        return d

    def _statusbar(self, *kwargs):
        """Called from various functions to set the statusbar text correctly depending on the set of conditions"""

        button = None
        if kwargs:
            # if kwargs exists, then it's an identifier for which control's corresponding function called it
            button = kwargs[0].GetEventObject()

        # Check to see if the src_file or src_dir has NOT been set (both empty strings)
        if not(self.src_file or self.src_dir):
            # set initial statusbar message
            self.status = "Begin by pressing the Select PDF button and select the PDF to split."

        # Check if the Select PDF button has been pressed but an output directory has not been set
        elif self.src_file and not self.out_dir:
            self.status = "Press 'Select Dir...' if you wish to change the output directory."

        # Finally, check to see if src_file has been set and src_dir has been selected
        elif self.src_file and self.out_dir:
            self.status = "Now confirm the output file prefix and suffix (requires alphanumeric) and press Apply."

    def _write_to_pdf(self, inputpdf):
        """Writes the split pdf files to the disk"""

        # lazy but tidy hack to set the value of out_prefix in one line instead of an if/else 4 liner
        out_prefix = self.txt_file_prefix.GetValue() or 'document-'

        # lazy but tidy hack to set the value of out_suffix in one line instead of an if/else
        out_suffix = self.txt_file_suffix.GetValue() or '001'

        out_dir = self.out_dir

        def _sufgen(suffix, length):
            yield(suffix)
            try:
                num = int(suffix)
                for x in range(num + 1, num + length):
                    yield x
            except ValueError:
                # set add (additional) to true initially because we've already yielded the first result
                add_one = True
                add_ten = False

                ord_suffix_dict = {123: 'a',  # 123 is the ordinal after 'z', 97 is ordinal for 'a'
                                    91: 'A',  # 91 is the ordinal after 'Z', 65 is ordinal for 'A'
                                    58: '0'}  # 58 is the ordinal after '9', 58 is ordinal for '0'
                # create a loop that iterates over a range of 'length - 1' size
                # (since we already yielded initial result)
                list_suffix = list(suffix)

                # run a loop that equals the maximum length the resulting string can be
                for _ in range(length - 1):
                    # loop over the string from right to left and begin iterating
                    for x in range(len(list_suffix) - 1, -1, -1):
                        # DEBUG ####
                        # print(f'CHAR: {list_suffix[x]}\nadd_one: {add_one}\nadd_ten: {add_ten}')
                        # DEBUG ####
                        o = ord(list_suffix[x])

                        if add_ten is True:
                            # if add_ten is True, we need to add one to the
                            # current ordinal of the character to the right
                            p = ord(list_suffix[x + 1]) + 1
                            # then we need to set that chr over to the 'reset value'
                            # i.e. '9' -> '0'; 'z' -> 'a'; 'Z' -> 'A'
                            list_suffix[x + 1] = ord_suffix_dict[p]

                            if o + 1 not in ord_suffix_dict:
                                add_one = True
                                add_ten = False

                        if add_one is True:
                            o += 1
                            if o not in ord_suffix_dict:
                                list_suffix[x] = chr(o)
                                break
                            else:
                                add_ten = True
                                add_one = False

                    yield ''.join(list_suffix)

        type_out_suffix = type(out_suffix)

        assert_error_msg = "Assert Error, out_suffix must be of type int or str"
        assert type_out_suffix in (str, int), self._dialog(assert_error_msg)

        if type_out_suffix is str and len(out_suffix) == 0:
            out_suffix = "001"

        # if the supplied variable is an integer, set len_out_suffix to the length of the integer
        if type_out_suffix is int:
            len_out_suffix = len(str(out_suffix))

        # elif the supplied string consists solely of integers, treat it like an integer
        elif all(map(lambda d: type(d) is int, out_suffix)):
            len_out_suffix = len(out_suffix)

        # else we must treat it like a string
        else:
            len_out_suffix = -1

        len_pdf = inputpdf.numPages

        # get the maximum length of the integer plus the number of pages to ensure the prefix is long enough for
        # all the pages, but if it's a string just use the length of the string and hope for the best (for now)
        len_max = max(len(str(int(out_suffix) + len_pdf)), len_out_suffix) if len_out_suffix != -1 else len(out_suffix)
        len_max = f'0{len_max}'

        suffix_gen = _sufgen(out_suffix, len_pdf)

        for i in range(len_pdf):
            # use the suffix generator to increment the suffix, using the maximum length obtained previously
            out_full_path = os.path.join(out_dir, f'{out_prefix}{next(suffix_gen):>{len_max}}.pdf')

            ##### DEBUG ######
            # print(out_dir, out_prefix, out_full_path)
            ##### DEBUG ######

            # Check for the existence of a file with the same name
            if not os.path.exists(out_full_path):
                try:
                    output = PdfFileWriter()
                    output.addPage(inputpdf.getPage(i))
                    with open(out_full_path, "wb") as outStream:
                        output.write(outStream)

                except OSError as ose:
                    err = f'OSError exception has occurred: \n{ose.args[0][1]}.\n\nContinue?'
                    dstyles = wx.YES_NO | wx.CENTRE | wx.STAY_ON_TOP
                    dlg = self._dialog(err, 'OSError', styles=dstyles)
                    if dlg.ShowModal() == wx.ID_YES:
                        dlg.Destroy()
                        continue
                    else:
                        dlg.Destroy()
                        break

                except Exception as e:
                    err = f'[-] Exception occurred during write: \n{e.args[0][1]}'
                    dstyles = wx.YES_NO | wx.CENTRE | wx.STAY_ON_TOP
                    dlg = self._dialog(err, 'OSError', styles=dstyles)
                    if dlg.ShowModal() == wx.ID_YES:
                        dlg.Destroy()
                        continue
                    else:
                        dlg.Destroy()
                        break
            else:
                # Name collision detected
                # TODO: Investigate "Yes to All" option
                msg = f'A file was found at "{out_full_path}", would you like to overwrite it?'
                dstyles = wx.YES_NO | wx.CANCEL | wx.CENTRE | wx.STAY_ON_TOP
                dlg = self._dialog(msg, "Overwrite existing file?", dstyles)
                result = dlg.ShowModal()
                if result == wx.ID_YES:
                    dlg.Destroy()
                    try:
                        output = PdfFileWriter()
                        output.addPage(inputpdf.getPage(i))
                        with open(out_full_path, "wb") as outStream:
                            output.write(outStream)

                    except OSError as ose:
                        err = f'OSError exception has occurred: \n{ose.args[0][1]}.\n\nContinue?'
                        dstyles = wx.YES_NO | wx.CENTRE | wx.STAY_ON_TOP
                        dlg = self._dialog(err, 'OSError', styles=dstyles)
                        if dlg.ShowModal() == wx.ID_YES:
                            dlg.Destroy()
                            continue
                        else:
                            break

                    except Exception as e:
                        err = f'[-] Exception occurred during write: \n{e.args[0][1]}'
                        dstyles = wx.YES_NO | wx.CENTRE | wx.STAY_ON_TOP
                        dlg = self._dialog(err, 'OSError', styles=dstyles)
                        if dlg.ShowModal() == wx.ID_YES:
                            dlg.Destroy()
                            continue
                        else:
                            break

                elif result == wx.ID_NO:
                    dlg.Destroy()
                    continue

                elif result == wx.ID_CANCEL:
                    dlg.Destroy()
                    break
        else:
            # Open explorer to output directory (Windows Only)
            os.startfile(out_dir)
            return True

    ####################################################################################################
    ##################################### EVENT CALLBACK FUNCTIONS #####################################

    def on_about(self, e):
        # Create a message dialog box
        msg = "About this software.\n\nThis program was created by Vic Jackson on behalf of Force Corporation. All " \
              "rights reserved. Copyright 2018. Intended solely for the use of Force Corporation employees. By " \
              "continuing to use the software, you hereby relinquish: any guarantee related to the software; any " \
              "technical assistance in seting up and/or using this software. \n\nIf you haven't gotten it yet, " \
              "this is private software and, as such, comes with zero guarantees, so please don't come asking.\n\nOn " \
              "the other hand, it's an extremely simple PDF splitter with step by step directions. Good luck!"

        dlg = wx.MessageDialog(self, msg, f"About PDF Splitter v{__version__}", wx.OK)
        dlg.ShowModal()  # Shows it
        dlg.Destroy()    # finally destroy it when finished.

    def on_exit(self, e):
        self.Close(True)  # Close the frame.

    def on_open(self, e):
        """ Runs when the 'Select PDF' button is pressed, allowing the user to choose a file to open"""
        dlg = wx.FileDialog(self.pnl, "Choose the PDF to split and save to individual pages", self.src_dir, "",
                            "PDF Files (*.pdf)|*.pdf", wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:  # If the OK button is clicked on the 'Select PDF' dialog
            self.src_file = dlg.GetFilename()
            self.src_dir = dlg.GetDirectory()

            self.txt_filename_preview.SetValue(os.path.join(self.src_dir, self.src_file))
            # First, check if the prefix is set or not. If it's not, set it to the prefix of the selected file
            if self.txt_file_prefix.GetValue() == '':
                self.txt_file_prefix.SetValue(self.src_file[:self.src_file.rfind('.')])

            # Now, check if the suffix txt box is filled in or not, if not, set it to the standard suffix
            if self.txt_file_suffix.GetValue() == '':
                self.txt_file_suffix.SetValue('001')

            # Next, check if the output directory has been selected
            if self.out_dir != '':
                oval = os.path.join(self.out_dir, self.txt_file_prefix.GetValue())
                suffix = f"{self.txt_file_suffix.GetValue()} [...]"
                if self.txt_file_suffix.GetValue() != '':
                    suffix = self.txt_file_suffix.GetValue()
                self.txt_outputprev.SetValue(f"{oval}{suffix}.pdf [...]")

            # Now, check if the output directory has not been selected
            elif self.txt_file_prefix.GetValue() and self.txt_file_suffix.GetValue() == '':
                # If not, use the base directory of the source file
                oval = f'{os.path.join(self.src_dir, self.txt_file_prefix.GetValue())}'
                self.txt_outputprev.SetValue(f"{oval}001.pdf [...]")
                self.out_dir = self.src_dir

            self._statusbar()

        dlg.Destroy()

    def on_select(self, e):
        """ Runs when output folder selected """
        style_dirdialog = wx.DD_DEFAULT_STYLE
        dlg = wx.DirDialog(self.pnl, "Choose the output folder...", self.src_dir, style=style_dirdialog)

        if dlg.ShowModal() == wx.ID_OK:
            self.out_dir = dlg.GetPath()

            file_prefix = self.txt_file_prefix.GetValue()
            file_suffix = self.txt_file_suffix.GetValue()

            if file_prefix and file_suffix:
                self.txt_outputprev.SetValue(f'{os.path.join(self.out_dir, file_prefix)}{file_suffix}.pdf [...]')
            elif file_prefix:
                self.txt_outputprev.SetValue(f'{os.path.join(self.out_dir, file_prefix)}001.pdf [...]')
            elif file_suffix:
                self.txt_outputprev.SetValue(f'{os.path.join(self.out_dir, self.src_file[:self.src_file.rindex(".")])}'
                                             f'{file_suffix}.pdf [...]')
            else:
                self.txt_outputprev.SetValue(f'{os.path.join(self.out_dir, self.src_file[:self.src_file.rindex(".")])}'
                                             '001.pdf')

            # update the statusbar based on the current conditions
            self._statusbar()
            # f.close()
        dlg.Destroy()

    def on_process(self, e):
        """ Process PDF and write separated PDF files to selected output directory"""
        btn = e.GetEventObject()
        btn.Disable()

        # ###### DEBUG ######
        # print(f'src_dir: {self.src_dir}\nsrc_file: {self.src_file}\nout_dir: '
        #       f'{self.out_dir}\nout_suffix: {self.txt_file_suffix.GetValue()}\n'
        #       f'out_suffix_type: {type(self.txt_file_suffix.GetValue())}')
        # ###### DEBUG ######

        self.txt_file_prefix.SetValue(self._fnslug(self.txt_file_prefix.GetValue()))
        src_fullpath = os.path.join(self.src_dir, self.src_file)
        src_pdf = PdfFileReader(open(src_fullpath, "rb"))
        result = self._write_to_pdf(src_pdf)

        if result:
            msg_success = "Processing completed successfully."
            dialog = self._dialog(msg_success, "Success")
            dialog.ShowModal()
            dialog.Destroy()
            self.btn_process.Enable()

        else:
            msg_fail = "Processing had an error finishing or was canceled."
            dialog = self._dialog(msg_fail, "Woops! Something went wrong.")
            dialog.ShowModal()
            dialog.Destroy()
            self.btn_process.Enable()

def main():

    app = wx.App(False)
    MainWindow(None, "ForcePDF Split and Save v1")
    app.MainLoop()


if __name__ == '__main__':
    main()
