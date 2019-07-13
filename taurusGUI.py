import tkinter as TK
import tkinter.scrolledtext as TKST
import tkinter.filedialog as TKFD

#########################################################################################################
#
#
#
#  CLASS TAURUS GUI
#
#
#
#########################################################################################################

class TaurusGUI(TK.Tk):

    IMAGEPATH       =   'img/'
    LOGOIMAGEFILE   =   IMAGEPATH + 'taurus.gif'    # classname not usable until class definition complete
    LABELFONT       =   ('Arial', 14, 'bold')
    LOGCAPFONT      =   ('Arial', 10, 'bold')
    TABCAPFONT      =   ('Arial', 10, 'bold italic')
    STRAPTEXT       =   'Taking the pain out of post-18'
    STRAPFONT        =   ('Roboto', 14, 'italic')

    def __init__(self, app, *args, **kwargs):
        TK.Tk.__init__(self, *args, **kwargs)
        self.app = app # for this class, the parent is the TaurusApp object
        self.wm_title("TAURUS - tracking and analysis unifying results, UCAS and SIMS data")
        self.geometry('800x600')
        self.option_add('*background', 'white')
        self.header = Header(self.app, self)
        self.logwin = LogWindow(self.app, self)
        self.mainwin = MainWindow(self.app, self)
        self.header.pack(side=TK.TOP, expand=False, fill=TK.X)
        self.mainwin.pack(side=TK.TOP, expand=True, fill=TK.BOTH)
        self.logwin.pack(side=TK.BOTTOM, expand=True, anchor=TK.N, fill=TK.BOTH)

    def end(self):      # called from app class when exit chosen
        self.destroy()

    def start(self):    # called from app class run() method
        self.refreshData()
        self.mainloop()

    def refreshData(self):
        self.header.refreshData()
        self.mainwin.refreshData()
        self.update_idletasks()

    def navigateTo(self, screen):
        self.mainwin.select(screen)

    def getBrowseLayout(self):  # required to return the BrowseLayout object to the application
        return self.mainwin.getBrowseLayout()

    def fileOpenDialog(self, **opts):
        return str(TKFD.askopenfilename(**opts))

    def fileSaveAsDialog(self, **opts):
        return str(TKFD.asksaveasfilename(**opts))

    def warning(self, message):
        message = self.app.preprocesswarning(message)
        self.logwin.log.configure(state=TK.NORMAL)
        self.logwin.log.insert(TK.END, message + '\n')
        self.logwin.log.yview(TK.END)                       # autoscroll to bottom
        self.logwin.log.configure(state=TK.DISABLED)

###############################################################################################################
#
#  CLASS HEADER and LOGWINDOW and MAINWINDOW
#
#  Basic layout of the GUI
#
###############################################################################################################

class Header(TK.Frame):

    def __init__(self, app, parent, *args, **kwargs):
        TK.Frame.__init__(self, parent, *args, **kwargs)
        self.app = app
        self.parent = parent
        self.labelwidth = 40  # in chars
        self.labeldata = None
        self.updateLabelData()
        self.makeWidgets()

    def getLabelData(self):
        return self.labeldata

    def updateLabelData(self):
        self.labeldata = self.app.getGUIlabeldata()

    def refreshData(self):
        self.updateLabelData()
        for i, (k, v) in enumerate(self.getLabelData()):
            self.textvars[i].set(self.prepareLabelText(k, v))

    def prepareLabelText(self, k, v):
        text = k + ': '
        if v is None: v = 'None'
        if len(v) + len(text) > self.labelwidth:
            c = self.labelwidth - len(text) - 3
            text += v[:c//2] + '...' + v[-c//2:]
        else:
            text += v
        return text

    def makeWidgets(self):        # Heading frame
        self.logoimage = TK.PhotoImage(file=TaurusGUI.LOGOIMAGEFILE)
        self.logolabel = TK.Label(self, image=self.logoimage)
        self.settings = TK.Frame(self)
        self.labels = []
        self.textvars = []
        for i in range(len(self.getLabelData())):
            self.textvars.append(TK.StringVar())
            self.labels.append(TK.Label(self.settings, textvariable=self.textvars[i], width=self.labelwidth, anchor=TK.W))
            self.labels[i].pack(fill=TK.BOTH, expand=True)
        self.refreshData()
        self.straplabel = TK.Label(self, text=TaurusGUI.STRAPTEXT, font=TaurusGUI.STRAPFONT)
        # future enhancement - grid these three using initial width and column weight
        # then the final column can resize but stay in proportion and the ... in the string can vary/disappear
        self.logolabel.pack(side=TK.LEFT)
        self.straplabel.pack(side=TK.LEFT, expand=1, anchor=TK.W)
        self.settings.pack(side=TK.RIGHT)

class LogWindow(TK.Frame):

    def __init__(self, app, parent, *args, **kwargs):
        TK.Frame.__init__(self,parent,*args,**kwargs)
        self.app = app
        self.parent = parent
        self.makeWidgets()

    def makeWidgets(self):
        self.logwincaption = TK.Label(self, text='Activity Log', font=TaurusGUI.LOGCAPFONT)
        self.logwincaption.pack(side=TK.TOP, expand=False, pady=(20,0))
        self.log = TKST.ScrolledText(self, height=6, state=TK.DISABLED)   # height is in textrows
        self.log.pack(side=TK.TOP, fill=TK.BOTH, expand=True)

class MainWindow(TK.Frame):

    def __init__(self, app, parent, *args, **kwargs):
        TK.Frame.__init__(self, parent, *args, **kwargs)
        self.app = app
        self.parent = parent
        self.layouts = {}
        self.active_layout = None
        self.makeWidgets()
        self.select('home')

    def refreshData(self):
        for layoutname, layoutframe in self.layouts.items():
            layoutframe.refreshData()

    def select(self, layout):
        if self.active_layout is not None:
            self.layouts[self.active_layout].pack_forget()
        self.layouts[layout].refreshData()
        self.layouts[layout].pack(fill=TK.BOTH, expand=True)
        self.active_layout = layout

    def getBrowseLayout(self):
        return self.layouts['browse']

    def makeWidgets(self):
        self.layouts['home'] = HomeLayout(self)
        self.layouts['home'].fillLayout()
        self.layouts['browse'] = BrowseLayout(self)
        self.layouts['browse'].fillLayout()
        self.layouts['report'] = ReportLayout(self)
        self.layouts['report'].fillLayout()
        self.layouts['manage'] = ManageLayout(self)
        self.layouts['manage'].fillLayout()

class Layout(TK.Frame):     # define interface

    def __init__(self):     # not used - overridden in subclasses
        pass

    def fillLayout(self):
        pass

    def refreshData(self):
        pass

###############################################################################################################
#
#  CLASS HOMELAYOUT
#
#  This is the homepage navigation - lives inside the MainWindow part of the application frame
#
###############################################################################################################

class HomeLayout(Layout):
    def __init__(self, parent, *args, **kwargs):
        TK.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.app = parent.app

    def fillLayout(self):
        self.importimage = TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'import.gif')
        self.analyseimage = TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'analyse.gif')
        self.settingsimage = TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'settings.gif')
        self.browseimage = TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'browse.gif')
        self.quitimage = TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'quit.gif')
        self.browse = TK.Button(self,image=self.browseimage,
                                command=lambda: self.app.navigateTo('browse'))
        self.report = TK.Button(self,image=self.analyseimage,
                                command=lambda: self.app.navigateTo('report'))
        self.manage = TK.Button(self,image=self.importimage,
                                command=lambda: self.app.navigateTo('manage'))
        self.config = TK.Button(self,image=self.settingsimage,
                                command=lambda: self.app.navigateTo('config'))
        self.quit = TK.Button(self,image=self.quitimage,command=lambda: self.app.quitApp())
        self.browselabel = TK.Label(self, text='Browse', font=TaurusGUI.LABELFONT)
        self.reportlabel = TK.Label(self, text='Report', font=TaurusGUI.LABELFONT)
        self.managelabel = TK.Label(self, text='Manage', font=TaurusGUI.LABELFONT)
        self.configlabel = TK.Label(self, text='Config', font=TaurusGUI.LABELFONT)
        self.quitlabel = TK.Label(self, text='Exit', font=TaurusGUI.LABELFONT)
        self.browse.grid(row=0, column=0, sticky=TK.S)
        self.report.grid(row=0, column=1, sticky=TK.S)
        self.manage.grid(row=0, column=2, sticky=TK.S)
        self.config.grid(row=0, column=3, sticky=TK.S)
        self.quit.grid(row=0, column=4, sticky=TK.S)
        self.browselabel.grid(row=1, column=0)
        self.reportlabel.grid(row=1, column=1)
        self.managelabel.grid(row=1, column=2)
        self.configlabel.grid(row=1, column=3)
        self.quitlabel.grid(row=1, column=4)
        for i in range(5):
            self.columnconfigure(i, weight=1)
        self.rowconfigure(0, pad=50)

###############################################################################################################
#
#  CLASS MANAGELAYOUT
#
#  Opened when user clicks manage from homepage - select a data source to import and analyse
#
###############################################################################################################

class ManageLayout(Layout):

    def __init__(self,parent,*args,**kwargs):
        TK.Frame.__init__(self,parent,*args,**kwargs)
        self.parent = parent
        self.app = parent.app

    def fillLayout(self):
        self.images = [ TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'asr.gif').zoom(2).subsample(3),
                        TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'simsMS.gif').zoom(2).subsample(3),
                        TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'simsCSV.gif').zoom(2).subsample(3),
                        TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'import.gif').zoom(2).subsample(3),
                        TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'import.gif').zoom(2).subsample(3),
                        TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'import.gif').zoom(2).subsample(3),
                        TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'quit.gif').zoom(2).subsample(3) ]
        self.callbacks = [ lambda: self.app.importASRdata(),
                           lambda: self.app.importSIMSpredictions(),
                           lambda: self.app.importFromSIMS(),
                           lambda: self.app.exportForSIMS(),
                           lambda: self.app.importBasedata(),
                           lambda: self.app.importResults(),
                           lambda: self.parent.select('home') ]
        self.buttons = []
        for i in range(7):
            self.buttons.append(TK.Button(self, image=self.images[i], command=self.callbacks[i]))
        self.labels = [ TK.Label(self, text='ASR Import', font=TaurusGUI.LABELFONT),
                        TK.Label(self, text='Marksheet Import', font=TaurusGUI.LABELFONT),
                        TK.Label(self, text='Report Import', font=TaurusGUI.LABELFONT),
                        TK.Label(self, text='Export UCAS', font=TaurusGUI.LABELFONT),
                        TK.Label(self, text='Basedata Import', font=TaurusGUI.LABELFONT),
                        TK.Label(self, text='Results Import', font=TaurusGUI.LABELFONT),
                        TK.Label(self, text='Back', font=TaurusGUI.LABELFONT) ]
        for i in range(7):
            self.buttons[i].grid(row=i//4*2, column=i%4)
            self.labels[i].grid(row=1+i//4*2, column=i%4)
            self.columnconfigure(i%4, weight=1)
        self.rowconfigure(0, pad=25, weight=1)
        self.rowconfigure(2, pad=25, weight=1)

###############################################################################################################
#
#  CLASS REPORTLAYOUT
#
#  Opened when user clicks reports from homepage - select a report to generate and save
#
###############################################################################################################

class ReportLayout(Layout):

    def __init__(self, parent, *args, **kwargs):
        TK.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.app = parent.app

    def fillLayout(self):
        self.images = [ TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'import.gif').zoom(2).subsample(3),
                        TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'import.gif').zoom(2).subsample(3),
                        TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'import.gif').zoom(2).subsample(3),
                        TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'import.gif').zoom(2).subsample(3),
                        TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'import.gif').zoom(2).subsample(3),
                        TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'import.gif').zoom(2).subsample(3),
                        TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'import.gif').zoom(2).subsample(3),
                        TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'quit.gif').zoom(2).subsample(3) ]
        self.callbacks = [ lambda: self.app.reportOffers(False),
                           lambda: self.app.reportOffers(True),
                           lambda: self.app.reportByStudent(),
                           lambda: self.app.reportByUni(),
                           lambda: self.app.reportDestinations(),
                           lambda: self.app.reportAtRisk(),
                           lambda: self.app.reportBySubject(),
                           lambda: self.parent.select('home') ]
        self.buttons = []
        for i in range(len(self.images)):
            self.buttons.append(TK.Button(self, image=self.images[i], command=self.callbacks[i]))
        self.labels = [ TK.Label(self, text='Status Updates', font=TaurusGUI.LABELFONT),
                        TK.Label(self, text='All Offers', font=TaurusGUI.LABELFONT),
                        TK.Label(self, text='By Student', font=TaurusGUI.LABELFONT),
                        TK.Label(self, text='By University', font=TaurusGUI.LABELFONT),
                        TK.Label(self, text='Destinations', font=TaurusGUI.LABELFONT),
                        TK.Label(self, text='At Risk', font=TaurusGUI.LABELFONT),
                        TK.Label(self, text='By Subject', font=TaurusGUI.LABELFONT),
                        TK.Label(self, text='Back', font=TaurusGUI.LABELFONT) ]
        for i in range(len(self.images)):
            if self.callbacks[i]:           # allow for blanks in grid
                self.buttons[i].grid(row=i//4*2, column=i%4)
            self.labels[i].grid(row=1+i//4*2, column=i%4)
            self.columnconfigure(i%4, weight=1)
        self.rowconfigure(0, pad=25, weight=1)
        self.rowconfigure(2, pad=25, weight=1)

###############################################################################################################
#
#  CLASS BROWSELAYOUT plus BROWSETOP, BROWSEBOTTOM and SEARCHBOTTOM
#
#  Opened when user clicks browse from homepage - browse students or show results of a search
#
###############################################################################################################

class BrowseLayout(Layout):

    def __init__(self, parent, *args, **kwargs):
        TK.Frame.__init__(self, parent, *args, **kwargs)
        self.rootwindow = parent.parent
        self.parent = parent
        self.app = parent.app
        self.studentname = TK.StringVar()
        self.dateshowing = TK.StringVar()

    def getGUIManager(self):
        return self.app.getGUIManager()

    def getApp(self):
        return self.parent.app

    def getSearchTerm(self):
        return self.topframe.getSearchTerm()

    def setStudentName(self, name):
        self.studentname.set(name)

    def setDateShowing(self, datestring):
        self.dateshowing.set(datestring)

    def setSearchTerm(self, value):
        return self.topframe.setSearchTerm(value)

    def fillLayout(self):
        self.topframe = BrowseTop(self)
        self.bottomframe = BrowseBottom(self)
        self.searchframe = SearchBottom(self)
        self.topframe.pack(side=TK.TOP, fill=TK.X, expand=False)
        self.choose('browse')

    def studentL(self):
        self.setStudentName(self.getGUIManager().decrementStudent())
        self.refreshData()

    def studentR(self):
        self.setStudentName(self.getGUIManager().incrementStudent())
        self.refreshData()

    def dateR(self):
        self.setDateShowing(self.getGUIManager().incrementDate())
        self.refreshData()

    def dateL(self):
        self.setDateShowing(self.getGUIManager().decrementDate())
        self.refreshData()

    def resetChosenDate(self):
        try:
            self.getGUIManager().resetDate()
            self.setDateShowing(self.getGUIManager().getFormattedDate())
        except AttributeError:  # browser object not instantiated yet
            self.setDateShowing('None')
        self.refreshData()

    def choose(self, flag):
        if flag == 'search':
            self.bottomframe.pack_forget()
            # srch.get(), do the searching then update the table widget in self.searchframe.tableGet()
            self.searchframe.pack(side=TK.TOP, fill=TK.BOTH, expand=True)
            self.searchframe.restart()      # reset search list to the top
        elif flag == 'browse' or True:      # ie all otherwise
            self.searchframe.pack_forget()
            self.resetChosenDate()
            self.bottomframe.pack(side=TK.TOP, fill=TK.BOTH, expand=True)
        self.refreshData()

    def clearSearch(self):
        self.setSearchTerm('')
        self.choose('browse')
        self.refreshData()

    def refreshData(self):
        try:
            self.studentname.set(self.getGUIManager().getStudent().getName())
        except AttributeError:     # not instantiated yet
            self.studentname.set('')
        try:
            self.dateshowing.set(self.getGUIManager().getFormattedDate())
        except AttributeError:     # not instantiated yet
            self.dateshowing.set('')
        self.bottomframe.refreshData()
        self.searchframe.refreshData()
        # reattach Enter key - is unbound when browse is closed
        self.rootwindow.bind("<Return>", self.topframe.enterPressed)

class BrowseTop(Layout):

    def __init__(self, parent, *args, **kwargs):
        TK.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        # choosebox contains the date and name pickers
        self.choosebox = TK.Frame(self)
        # name picker has arrows and name
        self.namebox = TK.Frame(self.choosebox)
        self.la1 = TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'left.gif').subsample(3)
        self.ra1 = TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'right.gif').subsample(3)
        self.lb1 = TK.Button(self.namebox, image=self.la1, command=lambda: self.parent.studentL())
        self.rb1 = TK.Button(self.namebox, image=self.ra1, command=lambda: self.parent.studentR())
        self.nm1 = TK.Label(self.namebox, textvariable=self.parent.studentname, width=36) # 36 chars
        self.lb1.pack(side=TK.LEFT)
        self.nm1.pack(side=TK.LEFT)
        self.rb1.pack(side=TK.LEFT)
        # date picker has arrows and date
        self.datebox = TK.Frame(self.choosebox)
        self.la2 = TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'left.gif').subsample(3)
        self.ra2 = TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'right.gif').subsample(3)
        self.lb2 = TK.Button(self.datebox, image=self.la2, command=lambda: self.parent.dateL())
        self.rb2 = TK.Button(self.datebox, image=self.ra2, command=lambda: self.parent.dateR())
        self.nm2 = TK.Label(self.datebox, textvariable=self.parent.dateshowing, width=36) # 36 chars
        self.lb2.pack(side=TK.LEFT)
        self.nm2.pack(side=TK.LEFT)
        self.rb2.pack(side=TK.LEFT)
        # date picker above name picker
        self.datebox.pack(side=TK.TOP)
        self.namebox.pack(side=TK.TOP)
        # end of choosebox
        # navbox contains the search entry field, search and cancel buttons and return to home button
        self.navbox = TK.Frame(self)
        self.backimg = TK.PhotoImage(file=TaurusGUI.IMAGEPATH+'quit.gif').subsample(3)
        self.back = TK.Button(self.navbox, image=self.backimg, command=self.close) # parent
        self.srch = TK.Entry(self.navbox)
        self.srchbut = TK.Button(self.navbox, text='SEARCH', anchor=TK.W,
                                 command=lambda: self.parent.choose('search'))
        self.srchclr = TK.Button(self.navbox, text='X', anchor=TK.W, command=lambda: self.parent.clearSearch())
        self.srch.pack(side=TK.LEFT, expand=True, fill=TK.X, anchor=TK.W)
        self.srchbut.pack(side=TK.LEFT, expand=False, fill=TK.X)
        self.srchclr.pack(side=TK.LEFT, expand=False, fill=TK.X)
        self.back.pack(side=TK.RIGHT, expand=True)
        # end of navbox
        # choosebox left of navbox
        self.choosebox.pack(side=TK.LEFT, fill=TK.BOTH, expand=True)
        self.navbox.pack(side=TK.RIGHT, fill=TK.BOTH, expand=True)
        # press return to search
        rootwindow = self.parent.rootwindow
        rootwindow.bind("<Return>", self.enterPressed)

    def close(self):
        mainwindow = self.parent.parent
        rootwindow = self.parent.rootwindow
        mainwindow.select('home')
        rootwindow.unbind("<Return>")

    def enterPressed(self, e):
        self.parent.choose('search')

    def getSearchTerm(self):
        return self.srch.get()

    def setSearchTerm(self, value):
        self.srch.delete(0, TK.END)
        self.srch.insert(0, value)

class BrowseBottom(Layout):

    MAXROWS = 8
    MAXRESULTS = 7

    def __init__(self, parent, *args, **kwargs):
        TK.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.bottomframeL = TK.Frame(self)
        self.personallabel = TK.Label(self.bottomframeL, text='Student Details', font=TaurusGUI.TABCAPFONT)
        self.bottomframeR = TK.Frame(self)
        self.bottomframeRT = TK.Frame(self.bottomframeR)
        self.bottomframeRB = TK.Frame(self.bottomframeR)
        self.choicelabel = TK.Label(self.bottomframeRT, text='Student Choices', font=TaurusGUI.TABCAPFONT)
        self.resulttablelabel = TK.Label(self.bottomframeRB, text='Student Results/Predictions', font=TaurusGUI.TABCAPFONT)
        # Collect data headings from application object
        self.personalheads, self.choiceheads, self.resultheads = self.getApp().getBrowseHeadings()
        self.personalwidth, self.choicewidth, self.resultwidth = self.getApp().getBrowseWidths()
        self.personaljusts, self.choicejusts, self.resultjusts = self.getApp().getBrowseJustifys()
        # Make the personal details panel (left side) - row layout so headings <> columns
        assert len(self.personalwidth) == len(self.personaljusts),\
            "Different number of headings and widths in BrowseBottomRT"
        self.personal = SimpleTable(self.bottomframeL, len(self.personalheads), 2,
                                    self.personalwidth, self.personaljusts, False)
        self.personallabel.pack(side=TK.TOP, fill=TK.X, pady=(10,20))
        self.personal.pack(side=TK.TOP, fill=TK.X)
        # Make the choices table (right top)
        assert len(self.choicewidth) == len(self.choiceheads) == len(self.choicejusts),\
            "Different number of headings and widths in BrowseBottomRT"
        self.choice = SimpleTable(self.bottomframeRT, BrowseBottom.MAXROWS,
                                  len(self.choiceheads), self.choicewidth, self.choicejusts, True)
        self.choicelabel.pack(side=TK.TOP, fill=TK.X, pady=(10,20))
        self.choice.pack(side=TK.TOP, fill=TK.X)
        # Make the results table (right bottom)
        widths = [15, 30, 20, 20, 20]
        assert len(self.resultwidth) == len(self.resultheads) == len(self.resultjusts),\
            "Different number of headings and widths in BrowseBottomRB"
        justifyleft = [False for i in range(len(self.resultheads))]
        justifyleft[1] = True   # subject name
        self.resulttable = SimpleTable(self.bottomframeRB, BrowseBottom.MAXRESULTS,
                                  len(self.resultheads), self.resultwidth, self.resultjusts, True)
        self.resulttablelabel.pack(side=TK.TOP, fill=TK.X, pady=(10,20))
        self.resulttable.pack(side=TK.TOP, fill=TK.X)
        # Pack right-side frames one above the other
        self.bottomframeRT.pack(side=TK.TOP, expand=True, fill=TK.BOTH)
        self.bottomframeRB.pack(side=TK.TOP, expand=True, fill=TK.BOTH)
        # Grid left and right frames adjacent - this enables the left side to have minimum size
        self.bottomframeL.grid(row=0,column=0,pady=10,padx=20,sticky='nsew')
        self.bottomframeR.grid(row=0,column=1,pady=10,sticky='nsew')
        self.columnconfigure(0,weight=0,pad=150) # don't resize this column
        self.columnconfigure(1,weight=1)

    def getApp(self):
        return self.parent.getApp()

    def getGUIManager(self):
        return self.getApp().getGUIManager()

    def refreshData(self):
        # Personal details
        appbrowser = self.parent.getApp().getGUIManager()
        if self.getGUIManager() is None:  # not instantiated yet
            return
        d = self.getGUIManager().getPersonalData()
        self.personal.populateV(self.personalheads, d)
        # Choices details
        self.choice.clearH(self.choiceheads, BrowseBottom.MAXROWS)
        d = self.getGUIManager().getChoiceData()
        self.choice.populateH(d)
        # Results details
        self.resulttable.clearH(self.resultheads, BrowseBottom.MAXRESULTS)
        d = self.getGUIManager().getResultData()
        self.resulttable.populateH(d)

class SearchBottom(Layout):

    def __init__(self, parent, *args, **kwargs):
        TK.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.tablerows = 11     # includes heading rows and MORE button
        self.tablecols = 7
        # Collect data headings from application object
        self.headings, self.widths, self.justify = self.parent.getApp().getSearchTableSettings()
        assert len(self.widths) == len(self.headings) == len(self.justify) == self.tablecols,\
            "Inconsistent number of table settings received from app in SearchBottom"
        self.table = ClickTable(self, self.tablerows, self.tablecols,
                                self.widths, self.headings, self.justify, True)
        self.table.pack(side=TK.TOP, fill=TK.X, pady=20)
        self.startat = 0

    def getGUIManager(self):
        return self.parent.getGUIManager()

    def cmd(self, i):
        if i >= 0:      # negative i is null
            self.getGUIManager().setStudentIndex(i)
            self.parent.choose('bottom')
            self.startat = 0    # next search starts at top

    def refreshData(self, sortcolumn=0):
        # clear table setting special cells appropriately
        studentID = 0 if self.getGUIManager() is None else self.getGUIManager().getStudentIndex()
        self.table.clear(self.cmd, studentID, '[Exit Search]')
        # paint data onto table widgets
        if self.getGUIManager():
            dataset = self.getGUIManager().search(self.parent.getSearchTerm().lower(), sortcolumn)
            for i in range(self.startat, min(self.startat+self.tablerows-2, len(dataset))):
                # lambda closure see https://docs.python.org/3/faq/programming.html
                # #why-do-lambdas-defined-in-a-loop-with-different-values-all-return-the-same-result
                self.table.setbutton(i-self.startat+1, dataset[i][1],
                                     lambda u=dataset[i][0]: self.cmd(u))
                for j in range(1, self.tablecols):  # tablecols excludes the name
                    self.table.set(i-self.startat+1, j, dataset[i][j+1])
            # set navigation widgets in table
            if self.startat+self.tablerows-2 >= len(dataset):
                self.table.setbutton(self.tablerows-1, '[Top]', lambda: self.tableTop(sortcolumn))
            else:
                self.table.setbutton(self.tablerows-1, '[More]', lambda: self.tableMore(sortcolumn))

    def tableMore(self, column):
        self.startat += self.tablerows-1
        self.refreshData(column)

    def tableTop(self, column):
        self.startat = 0
        self.refreshData(column)

    def tableSort(self, column):
        self.refreshData(column)

    def restart(self):
        self.startat = 0

###############################################################################################################
#
#  CLASS SIMPLETABLE
#
#  Simple tabular data - from Bryan Oakley
#
###############################################################################################################

class SimpleTable(TK.Frame):

    def __init__(self, parent, rows, columns, widths, justifies, highlighttop):
        TK.Frame.__init__(self, parent)
        self.parent = parent
        self._widgets = []
        for row in range(rows):
            current_row = []
            for column in range(columns):
                widget = self.makewidget(row, column, widths[column],
                                         justifies[column], highlighttop if row==0 else False)
                current_row.append(widget)
            self._widgets.append(current_row)
        # set all cols to resize by proportional width desired
        for column in range(columns):
            self.columnconfigure(column, weight=widths[column])

    def makewidget(self, row, column, width, justify, bold):
        font = ('Arial', 8) if not bold else ('Arial', 8, 'bold')
        anchor = TK.W if justify else TK.CENTER
        label = TK.Label(self, borderwidth=0, width=width, font=font, anchor=anchor)
        label.grid(row=row, column=column, sticky="nsew")
        return label

    def set(self, row, column, value, **kwargs):
        widget = self._widgets[row][column]
        widget.configure(text=value, **kwargs)

    def populateV(self, headings, data):
        for i, v in enumerate(headings):
            self.set(i, 0, v+':', anchor=TK.E, padx=10)
            self.set(i, 1, data[i], anchor=TK.W, padx=10)

    def populateH(self, data):
        for i, record in enumerate(data):
            for j, v in enumerate(record):
                self.set(i+1, j, v)

    def clearH(self, headings, maxrows):
        for i, v in enumerate(headings):
            self.set(0, i, v)
            for j in range(maxrows-1):  # clear rows except heading
                self.set(j+1, i, '')

###############################################################################################################
#
#  CLASS CLICKTABLE
#
#  Extension of simple table where first column and row are buttons (clickable)
#
###############################################################################################################

class ClickTable(SimpleTable):

    def __init__(self, parent, rows, columns, widths, headings, justifies, highlighttop):
        self.rows = rows
        self.cols = columns
        super().__init__(parent, rows, columns, widths, justifies, highlighttop)
        for i,title in enumerate(headings):
            self.setheading(i, title)

    def makewidget(self, row, column, width, justify, bold):
        if column == 0 or row == 0:
            # make button
            font = ('Arial', 8) if not bold else ('Arial', 8, 'bold')
            anchor = TK.W if justify else TK.CENTER
            button = TK.Button(self, borderwidth=0, width=width, font=font, anchor=anchor)
            button.grid(row=row, column=column, sticky="nsew")
            return button
        else:
            # makelabel as before
            return super().makewidget(row, column, width, justify, bold)

    def setbutton(self, row, text, command):
        widget = self._widgets[row][0]
        widget.configure(text=text, command=command)

    def setheading(self, column, text):
        widget = self._widgets[0][column]
        widget.configure(text=text, command=lambda: self.tablesort(column))

    def set(self, row, column, value, **kwargs):
        if row == 0 or column == 0:
            raise IndexError("Attempt to set heading/button in ClickTable")
        else:
            widget = self._widgets[row][column]
            widget.configure(text=value, **kwargs)

    def tablesort(self, column):
        self.parent.tableSort(column)

    def clear(self, callback, parameter=-1, toprow=''):
        # row 0 contains headings and doesn't get cleared
        for i in range(1, self.rows):
            if i == 1:
                self.setbutton(i, toprow, lambda: callback(parameter))
            else:
                self.setbutton(i, '', lambda: callback(-1))
            for j in range(1, self.cols):
                self.set(i, j, '')

############################################################################################
#
#   For testing only - file should be imported
#
############################################################################################

if __name__ == "__main__":
    app = None
    gui = TaurusGUI(app)
    gui.mainloop()
