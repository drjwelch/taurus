#########################################################################################################
#
#  TAURUS - tracking and analysis using results, UCAS and SIMS data
#           taking the pain out of post 18
#
#########################################################################################################

import taurusGUI
import datetime
import pickle
import os
import sys
import hashlib
import xml.sax as SAX
import xml.sax.saxutils as SAXUTILS

# Global constants

GLOBAL_ASR_DOB_FORMAT = '%d-%b-%y' # change to %Y if they go to 4 digit year in 2018 (millennials)
DEBUG   = True
OKAY    = 'Success' # global constant to return success from a method in another class - trying to retire this
MISSING = 'Missing' # global constant to indicate missing attribute value on another class - likewise retire

class JCQ:

    # Constants come from http://www.jcq.org.uk/exams-office/entries/jcq-formats/
    # jcq-formats-for-the-exchange-of-examination-related-data-revised-may-2016

    RESULT_TYPE_MARK         = 'UM'
    RESULT_TYPE_GRADE        = '123'
    RESULT_TYPE_MARKANDGRADE = 'BC'
    RESULT_TYPE_ALL          = RESULT_TYPE_MARK + RESULT_TYPE_GRADE + RESULT_TYPE_MARKANDGRADE

    # Recognise qual types
    SUBJECT_ALEVEL           = 'GCE A'
    SUBJECT_ASCERT           = 'GCE ASB' # note also ASD for double
    SUBJECT_EPQ              = 'EXPJB'
    SUBJECT_GCSEFC           = 'GCSEFC'
    SUBJECT_GCSESC           = 'GCSESC'
    SUBJECT_GCSE             = [SUBJECT_GCSEFC, SUBJECT_GCSESC]
    QUALEVELS                = [SUBJECT_ALEVEL, SUBJECT_ASCERT, SUBJECT_EPQ, SUBJECT_GCSE]

    # non-academic offer conditions ignorable in grade evaluation - count as if unconditional
    # In 2016 these were GMNX and non-A level codes were GKNMHL+ (plus prefix 1 for STEP - see Offer constructor)
    # In 2017, X removed as it is a GCSE condition so needs checking
    UCAS_IGNORE              = 'FGHLTJNM' # +XK are academic requirements
    # codes not relating to A level - stripped from Offer gradekey but left in full/rawgrade
    UCAS_NONALEVEL           = '+XKFGHLTJNM'

#########################################################################################################
#
#  CLASS STUDENT
#
#########################################################################################################

class Student():

    IDFIELDS = ['UPN', 'ULN', 'UCI', 'EXAMNO', 'NAMEHASH']

    def __init__(self, surname, forenames, dob, ucasID, cycle, pcode, idfields=None):
        self.surname = surname
        # multiple forenames not used but could use to distinguish students
        if ' ' in forenames:
            self.forename1 = forenames[:forenames.find(" ")]
            self.forename2 = forenames[forenames.find(" ")+1:]
        else:
            self.forename1 = forenames
            self.forename2 = ''
        self.fullname = self.surname.upper() + " " + self.forename1.capitalize()
        self.dob = datetime.datetime.strptime(dob,GLOBAL_ASR_DOB_FORMAT) # format is 02-DEC-99
        self.ucasID = ucasID
        self.cycle_no = cycle
        self.pcode = pcode
        self.choices = {} # key = date, value = list of Choice objects for this student
        # track interviews
        self.interviews = {} # key = choice ID, value = date
        self.isnew = True # set to false once we have an ASR history for this student
        # Following attributes are not from the ASR - they are imported from SIMS report export instead
        # Hash of fullname added to make checking by name easier later
        if idfields == None:
            idfields = [MISSING, MISSING, MISSING, MISSING]
        if len(idfields) == 4:
            idfields.append(hashlib.sha256(bytes(self.fullname,'utf-8')).hexdigest())
        self.ID = {}
        for i, field in enumerate(idfields):
            self.ID[Student.IDFIELDS[i]] = field
        # predicted grades imported from SIMS marksheet export
        self.predicted = {} # k=simsname, v=grade
        # tracking grades imported from SIMS marksheet export       NOT USED
        self.tracking = {}  # k=simsname, v=grade
        # populated on results day
        self.results = {}   # key = unit entry code; v= Result object
        self.isY13 = None   # identify students from UCAS who are in current Y13 cohort

    # Get methods for all the mandatory properties in the constructor

    def isNew(self):
        return self.isnew

    def isCurrentY13(self):
        return self.isY13   # initially set to None which equates to False in comparisons elsewhere

    def getUcasID(self):
        return self.ucasID

    def getName(self):
        return self.fullname

    def getSurname(self):
        return self.surname

    def getForename1(self):
        return self.forename1

    def getCycle(self):
        return self.cycle_no

    def getDOB(self):           # DOB is a datetime
        return self.dob

    def getDOBstring(self, fmt):
        return self.dob.strftime(fmt)

    def getPCode(self):
        return self.pcode

    def getChoices(self, thedate):
        if thedate not in self.choices:
            return []       # so still iterable if no choices, but empty
        else:
            return self.choices[thedate]

    def getChoicebyID(self, choiceID, theDate):
        if theDate in self.choices:
            for c in self.choices[theDate]:
                if c.getID() == choiceID:
                    return c
        return None

    # Get / set for optional constructor properties

    def getUCI(self):
        return self.getID('UCI')

    def getULN(self):
        return self.getID('ULN')

    def getUPN(self):
        return self.getID('UPN')

    def getExamNo(self):
        return self.getID('EXAMNO')

    def getID(self, identifier):
        if identifier not in self.ID:
            raise RuntimeError('Call to getID with invalid identifier '+identifier)
        if self.ID[identifier] is None:
            # can't currently fix this as too many places will fail if None returned
            return MISSING        #  ERRRRRRRRRRRRRRRR do we really want a str here?
        else:
            return self.ID[identifier]

    def setID(self, identifier, value):
        if self.validateID(identifier, value):
            self.ID[identifier] = value
        else:
            logwrite('@'+identifier+' '+value+" didn't validate")
            return False

    def setUPN(self, value):
        self.setID('UPN', value)

    def setUCI(self, value):
        self.setID('UCI', value)

    def setULN(self, value):
        self.setID('ULN', value)

    def setExamNo(self, value):
        self.setID('EXAMNO', value)

    def setYear(self, bool):
        self.isY13 = bool

    def validateID(self, identifier, value):
        char = value[0] if identifier == 'UPN' else value[-1]
        rest = value[1:] if identifier == 'UPN' else value[:-1]
        if identifier in ('EXAMNO', 'ULN'):
            try:
                test = int(value)
            except ValueError:
                return False
            if identifier == 'ULN' and (test >= 1e10 or test < 1e9):
                return False
            elif identifier == 'EXAMNO' and (test > 9999 or test < 0):
                return False
            else:
                return True
        elif identifier in ('UPN', 'UCI'):
            if not char.isupper():
                return False
            try:
                test = int(rest)
            except:
                return False
            if test >= 1e12 or test < 1e11:
                return False
            else:
                return True
        else:
            return False  # bad identifier
            
    def setNotNew(self):
        self.isnew = False

    #####################################################################################
    #
    # Main class methods for STUDENT
    #
    #####################################################################################

    def __eq__(self, other):
        try:  # comparing student objects
            if self.getUcasID() == other.getUcasID() \
                    or (self.getUCI() != MISSING and self.getUCI() == other.getUCI()) \
                    or (self.getUPN() != MISSING and self.getUPN() == other.getUPN()) \
                    or (self.getULN() != MISSING and self.getULN() == other.getULN()):
                return True
            else:
                return False
        except AttributeError: # in case type of other is not Student so has no getUcasID method
            return False

    def __hash__(self):
        return hash(self.ucasID)  # the only ID guaranteed to be present as created from ASR

    def __str__(self):
        return self.getName()

    def addChoice(self, thedate, choice):
        if thedate in self.choices:
            if choice not in self.choices[thedate]:     # eg if re-importing an asr ?????????
                self.choices[thedate].append(choice)
        else:                                           # first choice for this date
            self.choices[thedate] = [choice]
        return choice

    def countChoices(self, thedate, target, exact):
        if thedate not in self.choices:     # in case student has no choices at this date
            return 0
        if exact:
            return sum([1 for i in self.choices[thedate] if i.outcome == target])
        else:
            return sum([1 for i in self.choices[thedate] if i.outcome[:len(target)] == target])

    def getFirm(self, thedate):
        if thedate in self.choices:         # in case student has no choices at this date
            for choice in self.choices[thedate]:
                if choice.isFirm():
                    return choice
        return None

    def getInsc(self, thedate):
        if thedate in self.choices:         # in case student has no choices at this date
            for choice in self.choices[thedate]:
                if choice.isInsc():
                        return choice
        return None

    def addResult(self, result):
        unit = result.getUnitCode()
        if unit in self.results:
            return False        # duplicate grade in results: don't update
        else:
            self.results[unit] = result
        return True

    def getResults(self):
        return self.results

    def getResultbyUnit(self, unit):
        try:
            return self.results[unit]
        except KeyError:
            return None

    def getResultsAsOffer(self):
        # Return an Offer object made from grades that 'count' i.e. qual level results
        return Offer(''.join([result.getGrade() for code, result in self.results.items()]))

    def addPrediction(self, simsname, grade):
        # adds/updates predicted grade - returns old value if updated
        if simsname in self.predicted:
            temp = self.predicted[simsname]
            self.predicted[simsname] = grade  # always update even if already there
            if temp != grade:
                return temp     # return old value
            else:
                return False    # re-imported so not changed
        else:
            self.predicted[simsname] = grade
            return None         # no previous, so added

    def getPredictions(self):
        return self.predicted

    def getPredictionbySIMSName(self, simsname):
        if simsname in self.predicted:
            return self.predicted[simsname]
        return ''

    def getPredictionsAsOffer(self):
        g = ''
        for simsname in self.predicted:
            g += self.predicted[simsname]
        return Offer(g)

    def getPredictedGradeString(self, n=-1):
        o = self.getPredictionsAsOffer()
        r = o.getGrades(astar=True)
        del o
        if n == -1:
            n = len(r)
        return r[:n]

    def getPredictedSubjectCodeString(self):
        return ' '.join([simscode for simscode in self.predicted.keys()])

    def addInterview(self, choiceid, thedate):
        if choiceid in self.interviews:
            if datetime.datetime.strptime(self.interviews[choiceid],'%d%m%Y')\
                    < datetime.datetime.strptime(thedate,'%d%m%Y'):  # interviewed here earlier
                return False  # don't update
        # add date (or update a later dated INV to an earlier dated one we've just found)
        self.interviews[choiceid] = thedate
        return True

    def getInterviewDate(self, choiceid):
        if choiceid in self.interviews:
            return self.interviews[choiceid]
        else:
            return None

    def getUnconditionals(self, thedate):
        return self.countChoices(thedate,'U',Match.STARTSWITH)

    def getConditionals(self, thedate):
        return self.countChoices(thedate,'C',Match.STARTSWITH)

    def getDeclined(self, thedate):
        return self.countChoices(thedate,'CD',Match.STARTSWITH) + \
               self.countChoices(thedate,'UD',Match.STARTSWITH)

    def getInterviews(self, thedate):
        return self.countChoices(thedate,'INV',Match.EXACTMATCH)

    def getReferrals(self, thedate):
        return self.countChoices(thedate,'REF',Match.EXACTMATCH)

    def getRejections(self, thedate):
        return self.countChoices(thedate,'REJ',Match.EXACTMATCH)

    def getWithdrawals(self, thedate):
        return self.countChoices(thedate,'W',Match.EXACTMATCH)

    def getTotalChoices(self, thedate):
        return self.countChoices(thedate,'',Match.STARTSWITH)

    def getTotalOffers(self, thedate):
        return self.getUnconditionals(thedate) + self.getConditionals(thedate)

    def getOpenOffers(self, thedate):
        return self.getTotalOffers(thedate) - self.getDeclined(thedate)

    def getDecisions(self, thedate):
        return self.getTotalChoices(thedate) - self.getReferrals(thedate) - self.getInterviews(thedate)

    def getPossibleOffers(self, thedate):
        return self.getOpenOffers(thedate) - self.getReferrals(thedate)

    def acceptanceAnomaly(self, thedate):
        f = self.getFirm(thedate).getOffer()
        i = self.getInsc(thedate).getOffer()
        return f.getGradeValue() <= i.getGradeValue()

#########################################################################################################
#
#  CLASS CHOICE + related enumerated types
#
#########################################################################################################

class Update:
    UPD8_UNDEFINED  = -1
    UPD8_SAME       = 0
    UPD8_COURSE     = 1
    UPD8_OUTCOME    = 2
    UPD8_NEW        = 4

class Outcome:
    C   = 101
    U   = 102
    W   = 103
    REJ = 104
    INV = 105
    REF = 106

class Match:
    EXACTMATCH = True
    STARTSWITH = False

class Choice():

    def __init__(self, choiceid, unicode, unitext, crscode, crstext, outcome, offer):
        self.choiceid = choiceid
        self.unicode = unicode
        self.unitext = unitext
        self.crscode = crscode
        self.crstext = crstext
        self.outcome = outcome
        if self.getOutcome() == Outcome.U:
            self.offer = Offer('')                 # some unis leave grades even if U
        else:
            self.offer = Offer(offer.strip())       # grades in ASR often incl whitespace
        self.updated = Update.UPD8_UNDEFINED

    # Get methods

    def getID(self):
        return self.choiceid

    def getUni(self):       # unicode e.g. LEEDS is L23 (currently) is not used anywhere
        return self.unitext

    def getCrs(self):
        return self.crscode

    def getCrsText(self):
        return self.crstext
    
    def getOffer(self):
        return self.offer

    def getOfferGrades(self, astar=False):
        return self.getOffer().getGrades(astar)

    def getOfferGradeValue(self):
        return self.getOffer().getGradeValue()

    def getFullOutcome(self):
        return self.outcome

    def getOutcome(self):
        if self.outcome[0] == 'C':
            return Outcome.C
        elif self.outcome[0] == 'U':
            return Outcome.U
        elif self.outcome == 'REJ':
            return Outcome.REJ
        elif self.outcome == 'INV':
            return Outcome.INV
        elif self.outcome == 'REF':
            return Outcome.REF
        elif self.outcome[0] == 'W':
            return Outcome.W
        else:
            raise RuntimeError("getOutcome: invalid outcome "+self.outcome)  # need to handle any other cases?

    def getUpdated(self):
        return self.updated

    # Just one SET method needed

    def setUpdated(self, value):
        self.updated = value

    #####################################################################################
    #
    # Main class methods for CHOICE
    #
    #####################################################################################

    def __eq__(self, other):
        try:
            if self.choiceid == other.choiceid and self.unicode == other.unicode:
                return True
            else:
                return False
        except: # in case type of other is not Choice so has no choiceid/unicode method
            return False

    def __str__(self):
        raise NotImplementedError

    def hasUpdated(self):
        return self.updated != Update.UPD8_SAME     # what about undefined?

    def isInterview(self):
        return self.getOutcome() == Outcome.INV

    def isOffer(self):
        return self.getOffer().getFullGrades() != ''

    def setChoiceUpdateStatus(self, previousChoices):
        if self in previousChoices:     # equality of choices iff choicenum & uni match
            lastChoice = previousChoices[previousChoices.index(self)]
            self.setUpdated(Update.UPD8_SAME)
            if self.getCrsText() != lastChoice.getCrsText():
                self.setUpdated(self.getUpdated() | Update.UPD8_COURSE)
            if self.getOutcome() != lastChoice.getOutcome():
                self.setUpdated(self.getUpdated() | Update.UPD8_OUTCOME)
        else:
            self.setUpdated(Update.UPD8_NEW)

    def isFirm(self):
        if len(self.outcome) >= 2:
            if self.outcome[1] == 'F':
                return True
        return False

    def isInsc(self):
        if len(self.outcome) >= 2:
            if self.outcome[1] == 'I':
                return True
        return False

class Offer:
    
    MATCHED     =   0
    PLUS1       =   1
    PLUS2P      =   2
    MINUS1      =   3
    MINUS1P     =   4
    UP1DN1      =   5
    MIXED       =   6
    CHECK       =   7
    NOOFFER     =   8
    MET         =   [0, 1, 2]      # set of above conditions that count as 'met'
    UNMET       =   [3, 4, 5, 6]   # set of above conditions that count as 'unmet'
    DESCRIPTIONS =  ['Met', 'Above 1', 'Above 2+', 'Below 1', 'Below 2+',
                    'Up1Dn1', 'Mixed', 'CHECK', 'No Offer']
    SPECIALCONDITIONS = 1000    # sentinel for special academic offer

    def __init__(self, gradestring):
        # Keep copy of raw conditions with all whitespace gone and A*s not @s
        gradestring = gradestring.replace('@', 'A*')
        self.rawgrades = ''.join(gradestring.strip().split())

        # previous 1 prefix no longer used - all numeric values = points offer
        # self.rawgrades = self.rawgrades.lstrip('1')

        # Create key by removing conditions not interested in and sorting conditions
        # Amended Jul 2017 to reflect new UCAS codes

        # All non A level codes are stripped here
        grades = self.rawgrades.rstrip(JCQ.UCAS_NONALEVEL)
        # store A* internally as @ - all output of grades needs to replace back
        grades = grades.replace('A*', '@')
        try:
            v = int(grades)
            self.gradekey = grades                  # a string representing the integer offer
        except:
            self.gradekey = ''.join(sorted(grades)) # a string of grades in order

    def getFullGrades(self):
        return self.rawgrades

    def getGrades(self, astar=False):
        g = self.gradekey if not astar else self.gradekey.replace('@','A*')
        return g

    def numGrades(self):
        return len(self.gradekey)

    def isPointsOffer(self):
        try:
            value = int(self.getGrades())  # points offer
            return True
        except:
            return False

    def getGradeValue(self):        # turns offer grades into UCAS points value
        grades = self.getGradeEquivalent()
        if grades == '':    # no grade conditions
            if self.getFullGrades().rstrip(JCQ.UCAS_IGNORE) == '': # conditions that amount to U
                return 0    # this is effectively unconditional
            else:
                return Offer.SPECIALCONDITIONS      # ensure special academic conditions are considered unattainable
        value = 0
        for g in grades:
            try:
                value += self.gradeLettertoPoints(g)
            except ValueError:
                logwrite('#grade ' + g + ' not recognised - ignoring')
        return value

    def getGradeEquivalent(self, astar=False):   # turns a points offer into equivalent grades
        grades = self.getGrades(astar)
        if not self.isPointsOffer():
            return grades
        else:
            value = int(grades)
            trial = [5, 5, 5]             # start at three A*s
            trialvalue = 168
            while trialvalue > value:
                i = trial.index(max(trial))
                trial[i] -= 1           # decrease until we hit the right grade value
                trialvalue = sum([self.gradeNumbertoPoints(trial[i]) for i in range(3)])
            return ''.join([self.gradeLetterfromNumber(trial[i]) for i in range(3)])

    def gradeCompare(self, other, warn=True):
        # self is Offer object containing the result ('got') grades
        # other is also an Offer object containing the firm/insce offer or predictions
        # state values: 0 = matched, 1 = one above, 2 = >1 above,
        #               3 = 1 below, 4 = >1 below
        #               5 = +1-1, 6 = other mixed
        #               7 = CHECK (where extra conditions applied), 8 = No Offer

        # transition matrix for cases: g<o-1,g=o-1,g=o,g=o+1,g>o+1
        # matrix index is therefore sign(g-o) * min(abs(g-o),2) + 2
        # tmatrix[currentstate][case] = newstate
        tmatrix = [ [4,3,0,1,2], [6,5,1,2,2], [6,6,2,2,2], [4,4,3,5,6],
                    [4,4,4,6,6], [6,6,5,6,6], [6,6,6,6,6] ]
        # define the sign function
        sign = lambda x:(1,-1)[x<0]

        # Get string of grades (even for points offers)
        got = self.getGradeEquivalent()
        offer = other.getGradeEquivalent()

        # Handle special case
        if other.getGradeValue() == Offer.SPECIALCONDITIONS:
            return Offer.CHECK

        #start comparison state machine
        state = Offer.MATCHED
        for i in range(min(len(offer), len(got))): # only consider as many grades as needed by the offer condition
            try:
                g = self.gradeLettertoNumber(got[i])
                o = self.gradeLettertoNumber(offer[i])
            except ValueError:
                logwrite('#unrecognised grade: position ' + str(i) + ' result ' + got + ' offer ' + offer)
                return Offer.CHECK
            state = tmatrix[state][sign(g-o) * min(abs(g-o),2) + 2]
            if state == 5 and self.isPointsOffer():
                state = 0                       # up1dn1 equals met for a points offer
        logwrite('@result "' + got + '" offer "' + offer + '" outcome code ' + str(state))
        # optional warning if seemingly met but more conditions left unchecked
        if warn and state in Offer.MET and len(offer) > len(got):
            state = Offer.CHECK
        return state

    def gradeLettertoNumber(self, letter):
        return 'EDCBA@'.index(letter)

    def gradeLetterfromNumber(self, n):
        return 'EDCBA@'[n]

    def gradeNumbertoPoints(self, n):       # 2017 Tariff
        return n*8 + 16

    def gradeLettertoPoints(self, letter):
        return self.gradeNumbertoPoints(self.gradeLettertoNumber(letter))

#########################################################################################################
#
#  CLASS UNIRECORD
#
#########################################################################################################

class OfferType:

    UNCONDITIONALS = 0
    CONDITIONALS = 1
    REJECTIONS = 2

class UniRecord():

    def __init__(self,name):
        self.name = name
        self.outcomes = [0, 0, 0]   # unconditionals, conditionals, rejections
        self.offers = []            # list of offer objects

    # Get / set methods

    def getName(self):
        return self.name

    def getUnconditionals(self):
        return self.outcomes[OfferType.UNCONDITIONALS]

    def getConditionals(self):
        return self.outcomes[OfferType.CONDITIONALS]

    def getRejections(self):
        return self.outcomes[OfferType.REJECTIONS]

    def setUnconditionals(self, value):
        self.outcomes[OfferType.UNCONDITIONALS] = value

    def setConditionals(self, value):
        self.outcomes[OfferType.CONDITIONALS] = value

    def setRejections(self, value):
        self.outcomes[OfferType.REJECTIONS] = value

    #####################################################################################
    #
    # Main class methods for UNIRECORD
    #
    #####################################################################################

    def __str__(self):
        return self.name

    def __eq__(self, other):
        # can be called with other as a unirecord or with other as a string (the codename of a uni)
        # using str(other) in the comparison makes sure it works for both cases
        return self.name == str(other)

    def addOffer(self, offerstring):
        self.offers.append(Offer(offerstring))

    def getTotalOutcomes(self):
        return sum([self.outcomes[i] for i in range(3)])

    def getAllOffers(self):
        return self.offers

    def getNumOffers(self, grades):
        if 'A*' in grades: grades = grades.replace('A*','@')
        count = 0
        for offer in self.offers:
            if offer.getGrades() == grades:
                count += 1
        return count

#########################################################################################################
#
#  CLASS SUBJECT
#
#########################################################################################################

class Subject():

    def __init__(self, spec_code, unit_code, qual_level, unit_type, unit_name, max_ums=None, sims=''):
        self.speccode = spec_code   # spec code, not units e.g. '8704'
        self.unitcode = unit_code   # unit entry code e.q. PHYA1
        if qual_level.strip() == '' or qual_level not in JCQ.QUALEVELS:
            logwrite('#created subject but bad/empty qual level = ' + qual_level)
        self.qualevel = qual_level  # e.g. 'GCE AS'
        if unit_type not in JCQ.RESULT_TYPE_ALL:
            logwrite('#created subject but bad unit type = ' + unit_type)
        self.unittype = unit_type   # unit type e.g. unit, certification
        self.name     = unit_name   # e.g. 'AQA 2450 Physics ... etc.'
        self.simsname = sims        # e.g. Ph extracted from 'KS5 Ph UCAS Grade' - blank by default
        self.maxums   = max_ums     # maximum UMS for this unit     CURRENTLY NOT USED

    def __str__(self):
        return str(self.qualevel) + " (" + str(self.unittype)+") in " \
               + self.name + "(" + str(self.speccode) + "/" + str(self.unitcode) + ")"

    def getName(self):
        return self.name

    def getUnitCode(self):
        return self.unitcode

    def getUnitType(self):
        return self.unittype

    def getQualLevel(self):
        return self.qualevel

    def getSIMSName(self):
        return self.simsname

    def getMaxUMS(self):
        return self.maxums

    # Only these two settable after Subject is created from basedata import

    def setSIMSName(self, string):
        self.simsname = string

    def setMaxUMS(self, value):
        self.maxums = value

#########################################################################################################
#
#  CLASS RESULT and subclasses for type of result
#
#########################################################################################################

class Result():

    IDFIELDS = ['UCI', 'ULN', 'EXAMNO']

    def __init__(self, idfields, unitcode, grade, ums):
        self.unitcode = unitcode    # links to Subject object created by basedata
        self.grade = grade          # actual grade achieved
        self.ums = ums              # NOT USED CURRENTLY
        self.ID = {}
        if len(idfields) != len(Result.IDFIELDS):
            raise RuntimeError('Wrong idfields list passed to Result constructor')
        for i, field in enumerate(idfields):
            self.ID[Result.IDFIELDS[i]] = field

    def __str__(self):
        return 'Grade ' + self.grade + ' in ' + self.unitcode

    def getUCI(self):
        return self.ID['UCI']

    def getUnitCode(self):
        return self.unitcode

    def getGrade(self):
        # overridden in subclass for marks only
        if self.unitcode not in ['9000', '9001']:    # ignore AQA bacc
            return self.grade
        else:
            return ''

# Subclasses used to identify type of result

class ResultMark(Result):

    def __init__(self, idfields, unitcode, grade, ums):
        super().__init__(idfields, unitcode, grade, ums)

    def getGrade(self):
        # return blank string to represent no grade when result is a unit mark
        return '' # only the ResultGrade type is counted in results

class ResultGrade(Result):

    def __init__(self, idfields, unitcode, grade, ums):
        super().__init__(idfields, unitcode, grade, ums)

class ResultMarkGrade(Result):

    def __init__(self, idfields, unitcode, grade, ums):
        super().__init__(idfields, unitcode, grade, ums)

#########################################################################################################
#
#  CLASS DATASOURCE followed by subclasses BasedataDS, ResultDS and SIMSExtractDS
#
#########################################################################################################

class DataSource():
    # This class has the iter method implemented as a generator
    # This automatically returns an iterator object supplying the iter and next methods
    # See https://docs.python.org/3/library/stdtypes.html#typeiter 4.5.1
    def __init__(self, app, filepath, keyword, initial):
        self.app = app
        self.filepath = filepath
        self.keyword = keyword
        self.initial = initial

    def __iter__(self):
        filelist = self.getDSfilelist()
        if len(filelist) == 0:
            logwrite('No ' + self.keyword + ' files found')
            raise StopIteration
        for file in filelist:
            logwrite('#trying file '+file)
            datasource = self.app.trytoopen(os.path.join(self.filepath, file),
                                          'cannot open file %F: maybe already open?')
            if datasource == TaurusApp.OPENFAIL:
                raise StopIteration
            else:
                with datasource:
                    for line in datasource:
                        yield self.processLine(line)

    def getDSfilelist(self):
        return [filename for filename in os.listdir(self.filepath)
                if filename[0].upper() == self.initial and filename.upper()[-4:-2] == '.X']

    def processLine(self, line):
        pass                    # needs overriding in subclasses

#########################################################################################################
#
#  CLASS BASEDATADS
#
#########################################################################################################

class BasedataDS(DataSource):

    def __init__(self, app):
        super().__init__(app, app.getFullPath('EXAMSIN'), 'basedata', 'O')

    def processLine(self, line):
        if line[0:2] == 'O5':  # indicates unit line otherwise header
            unit_code = line[2:8].rstrip()
            spec_code = line[8:14].rstrip()
            qual_level = line[14:21].rstrip()
            unit_type = line[21].rstrip() # is C cert U unit B both
            unit_name = line[42:78].rstrip()
            try:
                max_ums = int('0'+line[109:113])
            except ValueError:
                max_ums = 0
                logwrite('#ignoring invalid max UMS value: ' + line[109:113])
            return Subject(spec_code, unit_code, qual_level, unit_type, unit_name, max_ums)
        else:
            logwrite('#skipping header line: ' + line)

#########################################################################################################
#
#  CLASS RESULTDS
#
#########################################################################################################

class ResultDS(DataSource):

    def __init__(self, app):
        super().__init__(app, app.getFullPath('EXAMSIN'), 'basedata', 'R')


    def processLine(self, line):
        if line[0:2] == 'R5':           # indicates result line otherwise header
            center_number = line[2:7]   # future use for license check
            exam_number = line[7:11]
            uci = line[11:24]
            uln = line[24:34]
            # columns 34-40 are spaces
            raw_result = line[40:]
            idfields = [uci, uln, exam_number]
            return self.createResult(idfields, raw_result)
        else:   # line not starting R5
            logwrite('#skipping header line: ' + line)
            return None

    def createResult(self, idfields, raw_result):
        unitcode = raw_result[0:6].rstrip()  # not int because it sometimes has letters
        unittype = raw_result[6]
        if unittype in JCQ.RESULT_TYPE_GRADE:
            grade = raw_result[7:9].rstrip()
            ums = None
            return ResultGrade(idfields, unitcode, grade, ums)
        elif unittype in JCQ.RESULT_TYPE_MARKANDGRADE:
            grade = raw_result[11:13].rstrip()
            ums = int(raw_result[7:11])
            return ResultMarkGrade(idfields, unitcode, grade, ums)
        elif unittype in JCQ.RESULT_TYPE_MARK:
            grade = raw_result[10:12]
            ums = int(raw_result[7:10])
            return ResultMark(idfields, unitcode, grade, ums)

#########################################################################################################
#
#  CLASS SIMSExtractDS
#
#########################################################################################################

class SIMSExtractDS(DataSource):

    FORENAMEFIELDS = ['LEGAL FORENAME', 'LFORENAME', 'FORENAME']
    SURNAMEFIELDS  = ['LEGAL SURNAME', 'LSURNAME', 'SURNAME']
    UPDATEFIELDS   = Student.IDFIELDS
    DOBKEYS        = ['DOB', 'DATE OF BIRTH']
    PCODEKEYS      = ['POSTCODE', 'PCODE']
    EXAMNOKEYS     = ['EXAM NUMBER', 'EXAMNO', 'EXAM NO']

    def __init__(self, app):
        # user chooses sims report csv file
        opts = {  'defaultextension': '.csv',
                  'initialdir': app.config['ASRPATH'],
                  'title': 'Choose SIMS Report file'    }
        chosenfile, self.file = os.path.split(app.chooseFiletoOpen(opts))
        # set it up as a DS
        super().__init__(app, chosenfile, 'chosen', None)
        self.d = {}             # temp dict to hold csv data
        self.headings = []      # list of headings which are keys to d

    def getDSfilelist(self):    # overridden to return only the chosen file
        return [self.file]

    def processLine(self, line):
        if self.headings == []:      # first line; extract headings
            self.setHeadings(line)
        else:
            self.prepareRecord(line)
            student = self.identifyStudent()
            if student is None:
                logwrite('ignoring student not found in UCAS: line begins ' + line[:20])
            return student

    def setHeadings(self, line):
        self.headings = line.strip().upper().split(',')
        for i, key in enumerate(self.headings):
            if key in SIMSExtractDS.DOBKEYS:
                self.headings[i] = key = 'DOB'
            if key in SIMSExtractDS.EXAMNOKEYS:
                self.headings[i] = key = 'EXAMNO'
            if key in SIMSExtractDS.PCODEKEYS:
                self.headings[i] = key = 'POSTCODE'
            self.d[key] = 0

    def prepareRecord(self, line):
        record = line.strip().split(',')
        for i in range(len(self.headings)):
            self.d[self.headings[i]] = record[i]
        if 'DOB' in self.d:
            self.d['DOB'] = datetime.datetime.strptime(self.d['DOB'],'%d %B %Y')

    def identifyStudent(self):
        studentmanager = self.app.getStudentManager()
        me = None
        if 'UPN' in self.d:     # try finding a UPN match first
            me = studentmanager.getStudentbyUPN(self.d['UPN'])
        if not me:
            # UCI ULN ExamNo may only be in results file and not loaded yet
            # so try matching on dob/pcode which is loaded early as it comes from UCAS ASR
            candidatesDOB = set([])     # may be >1 student with same DOB
            candidatesPCD = set([])     # likewise postcode
            if 'DOB' in self.d:
                candidatesDOB = set(studentmanager.getStudentbyDOB(self.d['DOB']))
            if 'POSTCODE' in self.d:
                candidatesPCD = set(studentmanager.getStudentbyPostcode(self.d['POSTCODE']))
            if len(candidatesDOB) + len(candidatesPCD) != 0:
                candidates = candidatesDOB.intersection(candidatesPCD)  # anyone matching both DOB and postcode?
                if len(candidates) == 1:      # one student matched both: must be the right one
                    me = list(candidates)[0]
                elif len(candidates) > 1:       # separate by using names ... full names as maybe twins!
                    for s in candidates:
                        me = self.matchByName(s)
                        if me: break
                elif len(candidates) == 0:
                    # No candidate matches both; check if a name matches any of those with EITHER pcode or DOB match
                    # If so, warn in case of dirty data in SIMS or ASR
                    for s in candidatesDOB.union(candidatesPCD):
                        me = self.matchByName(s)
                        if me: break
                    if me is not None:
                        logwrite('found UCAS student with matching name ' + me.getName() +
                                 ' but mismatch fields: SIMS import DOB ' + str(self.d['DOB']) +
                                 ' Postcode ' + self.d['POSTCODE'] +
                                 ', but UCAS DOB ' + me.getDOBstring('%d %m %Y') + ' Postcode ' + me.getPCode() )
                        logwrite('Student was imported anyway; review postcode and DOB data in SIMS')
        self.updateStudentRecord(me)
        # finally check exam number, if provided, is unique
        if me and ('EXAMNO' in self.d):
            duplicates = studentmanager.getStudentbyExamNo(me.getExamNo())
            if len(duplicates) > 1:
                msg = ", ".join([s.getName() for s in duplicates])
                logwrite('warning: duplicate exam numbers for: ' + msg)
                logwrite('resolve and re-import')
        return me

    def matchByName(self, s):
        forename_me = s.getForename1()
        forename_DS = self.getAnyAvailableForename()
        surname_me = s.getSurname()
        surname_DS = self.getAnyAvailableSurname()
        if surname_me == surname_DS and forename_me == forename_DS:
            return s
        else:
            return None

    def updateStudentRecord(self, student):
        if student is not None:
            for field in SIMSExtractDS.UPDATEFIELDS:
                if field in self.d:
                    student.setID(field, self.d[field])
            # postcode and dob from SIMS file are not used - UCAS ASR data is the master
            student.setExamNo(self.getAnyAvailableExamNumber())
            if 'Y13' in self.d:
                student.setYear(self.d['Y13']=='Year 13')

    def getAnyAvailableExamNumber(self):
        availableexamnofields = [field in self.d for field in SIMSExtractDS.EXAMNOKEYS]
        if any(availableexamnofields):
            fieldnumber = availableexamnofields.index(True)
            return self.d[SIMSExtractDS.EXAMNOKEYS[fieldnumber]]

    def getAnyAvailableSurname(self):
        availablesurnamefields = [field in self.d for field in SIMSExtractDS.SURNAMEFIELDS]
        if any(availablesurnamefields):
            fieldnumber = availablesurnamefields.index(True)
            return self.d[SIMSExtractDS.SURNAMEFIELDS[fieldnumber]]

    def getAnyAvailableForename(self):
        availableforenamefields = [field in self.d for field in SIMSExtractDS.FORENAMEFIELDS]
        if any(availableforenamefields):
            fieldnumber = availableforenamefields.index(True)
            return self.d[SIMSExtractDS.FORENAMEFIELDS[fieldnumber]]

#########################################################################################################
#
#  CLASS ASRFILE
#
#########################################################################################################

class ASRFile():

    def __init__(self, app, filename):
        self.app = app
        self.filename = filename
        self.handle = None

    def __enter__(self):
        self.handle = self.app.trytoopen(self.filename, 'failed to open ASR candidate file %F')
        if self.handle != TaurusApp.OPENFAIL:
            # Get file date from first line of file - format is 23/07/2016
            self.data = self.handle.readline().strip()
            self.unescape()
            self.data = self.data.split(",")
            self.filedate = datetime.datetime.strptime(self.data[0][42:52],"%d/%m/%Y")
            # Get establishment number from second line - estab no + estab name
            self.data = self.handle.readline().strip()
            self.unescape()
            self.data = self.data.split(",")
            self.estab = self.data[0]
        return self

    def getStatus(self):    # returns OPENFAIL if it failed
        return self.handle

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.handle.close()

    def __iter__(self):
        return self

    def __next__(self):
        self.data = self.handle.readline().strip()
        if self.data == "":  #EOF
            raise StopIteration
        else:
            self.unescape()
            self.data = self.data.split(",")
            return self.data

    def getEstabNo(self):
        return self.estab

    def getFileDate(self):
        return self.filedate

    def unescape(self):
        if "\"" in self.data:
            escapecommas = False
            c=0
            while c<len(self.data):
                if self.data[c] == "\"": # if string quoted, remove quote
                    escapecommas = not escapecommas
                    self.data = self.data[:c] + self.data[c+1:]
                    c-=1 # string has shrunk by 1
                elif self.data[c] == "," and escapecommas: # and replace contained commas with spaces
                    self.data = self.data[:c] + ' ' + self.data[c+1:]
                c+=1

    def isStudent(self):
        try: # DoB in field 3 and UCASID in 4
            d = datetime.datetime.strptime(self.data[2],GLOBAL_ASR_DOB_FORMAT)
            u = int(self.data[3])
        except (ValueError, IndexError): # if not int or date, or is a short line
            return False
        else:
            return (u>1E9 and u<1E10) # any 10 digit int is OK for a UCASID

    def getStudentFields(self):
        if self.isStudent():
            # fields needed for Student object constructor
            # surname, forenames, dob, UCASid, cycle num, pcode
            return (self.data[0],self.data[1],self.data[2],self.data[3],self.data[4],self.data[9])
        else:
            return None

    def isChoice(self):
        try: # choice has int in first field and year of uni entry in 9th
            i = int(self.data[0])
            y = int(self.data[8])
        except (ValueError, IndexError): # if not int or short line
            return False
        else:
            return (y>2000 and y<2100) # should cover all plausible years and reject most errors

    def getChoiceFields(self):
        if self.isChoice():
            # fields needed for Choice object constructor
            # choice number, unicode, unitext, crscode, crstext, outcome, offer (grades)
            return (self.data[0],self.data[1],self.data[2],self.data[3],
                    self.data[7],self.data[5],self.data[6])
        else:
            return None

#########################################################################################################
#
#  CLASS STUDENTMANAGER
#
#########################################################################################################

class StudentManager():

    def __init__(self, app):
        self.app = app
        self.students = []      # list of student objects
        self.withDates = []     # list of ASR dates used to assemble students list

    def __iter__(self):
        self.ptr = -1
        return self
            
    def __next__(self):
        self.ptr += 1
        if self.ptr == len(self.students):
            raise StopIteration
        return self.students[self.ptr]

    def isLoaded(self):
        return len(self.students)!=0

    def addStudent(self, historic, *student):
        newstudent = Student(*student)
        if newstudent not in self.students:
            insertpoint = 0
            for i, s in enumerate(self.students):
                insertpoint = i
                if newstudent.getName() < s.getName():
                    break
                insertpoint += 1  # if at end will fall out of loop, so append
            if historic:
                logwrite('warning: student in historic data is not in current data: ' + newstudent.getName())
            self.students.insert(insertpoint,newstudent)
        else:
            self.students[self.students.index(newstudent)].setNotNew()
        return self.students[self.students.index(newstudent)]

    def getCurrentDate(self):
        if len(self.withDates) == 0:
            logwrite('#attempt to get current date when none loaded')
            return None
        return self.withDates[0]

    def getPreviousDate(self):
        if len(self.withDates) <= 1:
            logwrite('#attempt to get previous date when none available')
            return None
        return self.withDates[1]

    def getAllDatesSeen(self):
        return self.withDates

    def addDate(self, thedate):
        if len(self.withDates) == 0:
            self.withDates.append(thedate)
        else:
            insertpoint = 0
            for i, d in enumerate(self.getAllDatesSeen()):
                insertpoint = i
                if datetime.datetime.strptime(thedate,'%d%m%Y') > datetime.datetime.strptime(d,'%d%m%Y'):
                    break
                insertpoint += 1  # in case of adding at end of list
            self.withDates.insert(insertpoint, thedate)

    def getStudentbyPosition(self,n):
        try:
            s = self.students[n]
            return s
        except IndexError:
            return None

    def getNumStudents(self):
        return len(self.students)

    def getStudentbyExamNo(self,target):
        r = []
        for s in self.students:
            if s.getExamNo() == target:
                r.append(s)
        return r

    def getStudentbyUPN(self, target):
        for s in self.students:
            if s.getUPN() == target:
                return s
        return None

    def getStudentbyUCI(self,target):
        for s in self.students:
            if s.getUCI() == target:
                return s
        return None

    def getStudentbyULN(self,target):
        for s in self.students:
            if s.getULN() == target:
                return s
        return None

    def getStudentbyDOB(self,target):
        r = []
        for s in self.students:
            if s.getDOB() == target:
                r.append(s)
        return r

    def getStudentbyPostcode(self,target):
        r = []
        for s in self.students:
            if s.getPCode() == target:
                r.append(s)
        return r

    def getStudentbySurname(self,target):
        r = []
        for s in self.students:
            if s.getSurname() == target:
                r.append(s)
        return r

    def getStudentfromResult(self, result):
        for s in self.students:
            if s.getUCI() == result.getUCI():  # both objects implement a getUCI method used for comparison
                return s
        return None

    def loadStudents(self):
        pklfile = self.getCurrentPKL()
        if pklfile is None:
            logwrite('failed as no data files available')
            return
        logwrite('try loading student data from ' + pklfile)
        # File was identified by opening it in getCurrentPKL so no need to check it will open
        with open(pklfile,'rb') as f:
            token = pickle.load(f)
            if token != TaurusApp.DATAFILETOKEN:
                raise RuntimeError('Loading students from file with no token')
            self.app.validateLicence(pickle.load(f),'student datafile '+pklfile)
            self.withDates = pickle.load(f)
            self.students = pickle.load(f)
        logwrite('success!')
        return True

    def getCurrentPKL(self):
        # Find which file in the PKLPATH directory is from the latest-dated ASR
        # Return latest filename
        # If no such file, value None returned
        path = self.app.getFullPath('PKLPATH')
        logwrite('try loading existing student data from ' + path)
        pklfiles = [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f)) and f[-4:] == '.pkl']
        if len(pklfiles)==0:
            logwrite('no data files found in ' + path)
            return None
        else:
            latestDate = datetime.datetime.strptime('01011900','%d%m%Y')
            latestFile = -1
            allDatesSeen = set()  # empty set
            for i, file in enumerate(pklfiles):
                f = self.app.trytoopen(path+file, 'skipping candidate pickle file %F - cannot open', mode='rb')
                if f == TaurusApp.OPENFAIL:
                    continue
                with f:
                    token = pickle.load(f)
                    if token != TaurusApp.DATAFILETOKEN:    # go to next file if not a student pickle
                        continue
                    self.app.validateLicence(pickle.load(f),'student datafile '+file)
                    filedates = pickle.load(f)
                    for thedate in filedates:
                        self.addDate(thedate)
                    allDatesSeen = allDatesSeen.union(set(self.getAllDatesSeen()))
                    thisDateString = self.getCurrentDate()
                    thisDate = datetime.datetime.strptime(thisDateString,'%d%m%Y')
                    if thisDate > latestDate:
                        latestDate = thisDate
                        latestFile = i
            if allDatesSeen != set(self.getAllDatesSeen()):
                logwrite('the latest data file does not include the ASR data from: ' +
                         str(allDatesSeen.difference(set(self.getAllDatesSeen()))))
                logwrite("you should import ASR data to collect and save the new files' data")
            if latestFile == -1:
                logwrite('seemingly no ASR file available dated >1900 so no student data was loaded')
            else:
                logwrite('success - newest data file includes ASR from ' + latestDate.strftime('%d%m%Y') +
                         ', filename ' + pklfiles[latestFile])
                return os.path.join(path, pklfiles[latestFile])

    def saveStudents(self):
        pklfile = self.getPickleFileName()
        f = self.app.trytoopen(pklfile, 'unable to open savefile %F', mode='wb')
        if not f or f == TaurusApp.OPENFAIL:
            logwrite('#student data could not be saved')
            return
        with f:
            pickle.dump(TaurusApp.DATAFILETOKEN, f)
            pickle.dump(self.app.getLicenseToken(), f)
            pickle.dump(self.withDates, f)
            pickle.dump(self.students, f)
        return pklfile

    def getPickleFileName(self):
        try:
            return self.app.getFullPath('PKLPATH')+self.app.getConfig('PKLNAME')+ \
               self.getCurrentDate()+'.pkl'
        except:
            return None

#########################################################################################################
#
#  CLASS SUBJECTMANAGER
#
#########################################################################################################

class SubjectManager():

    def __init__(self, app):
        self.app = app
        self.subjects = []      # list of subject objects

    def getSubjects(self):
        return self.subjects

    def getNumSubjects(self):
        return len(self.subjects)

    def saveSubjects(self):
        pklfile = self.getSubjectsFileName()
        f = self.app.trytoopen(pklfile, 'unable to save subject basedata to %F', mode='wb')
        if f == TaurusApp.OPENFAIL:
            return
        with f:
            pickle.dump(TaurusApp.BASEFILETOKEN, f)
            pickle.dump(self.app.getLicenseToken(), f)
            pickle.dump(self.subjects, f)

    def loadSubjects(self):
        pklfile = self.getSubjectsFileName()
        logwrite('try loading subject data from ' + pklfile)
        f = self.app.trytoopen(pklfile, 'unable to load subject basedata from %F - file not found?', mode='rb')
        if f == TaurusApp.OPENFAIL:
            return False
        with f:
            token = pickle.load(f)
            if token != TaurusApp.BASEFILETOKEN:
                logwrite('#unable to load subjects from this file as it has no token')
                return False
            self.app.validateLicence(pickle.load(f),'student datafile '+pklfile)
            self.subjects = pickle.load(f)
        logwrite('success!')
        return True

    def getSubjectsFileName(self):
        return self.app.getFullPath('PKLPATH')+'basedata.pkl'

    def getUnitCodesfromBasedata(self):
        return [s.getUnitCode() for s in self.subjects]

    def mapSubjects(self):
        csv_file = self.getMappingFileName()
        logwrite('try applying subject mappings from ' + csv_file)
        f = self.app.trytoopen(csv_file, 'file not found - skipping subject mapping')
        if f == TaurusApp.OPENFAIL:
            return
        read_scs = f.readline().strip().split(',')
        row = f.readline()
        while row != '':
            fields = row.strip().split(',')
            n = sum([1 for field in fields if field != ''])
            if n > 2:   # one for field name and one for the indicator mark
                logwrite('warning: >1 mapping for code ' + fields[0] + ': only left-most column is processed')
            for i, field in enumerate(fields):
                if field != '' and i != 0:
                    subject = self.getSubjectbyUnitCode(fields[0])
                    if not subject:
                        logwrite('#subject code in mapping is not in basedata: '+fields[0])
                    else:
                        logwrite('mapping ' + fields[0] + ' to subject code ' + read_scs[i])
                        subject.setSIMSName(read_scs[i])
                    break
            row = f.readline()
        f.close()
        logwrite('success - any mappings listed above were completed')

    def updateSubjectMapping(self):
        csv_file = self.getMappingFileName()
        logwrite('try writing subject mappings to ' + csv_file)
        unitcodes = self.getUnitCodesfromBasedata()
        simscodes = self.getSIMSCodesfromPredictions()
        # if file exists, read the top header and following lines
        infile = self.app.trytoopen(csv_file, "cannot read to update subject mappings: file doesn't exist or is open?")
        if infile == TaurusApp.OPENFAIL:
            read_scs = ['']
            rows = []
        else:
            read_scs = infile.readline().strip().split(',')
            rows = infile.readlines()
            infile.close()
        # write top header with any new simscodes added to the end
        outfile = self.app.trytoopen(csv_file, 'cannot write to update subject mappings: file open?', mode='w')
        if outfile == TaurusApp.OPENFAIL:
            return
        for c in simscodes:
            if c not in read_scs:
                read_scs.append(c)
        outfile.write(','.join(read_scs)+'\n')
        numcols = len(read_scs)
        # write existing rows and add any new unit codes as new rows at the bottom
        for row in rows:
            record = row.strip().split(',')
            if len(record) != numcols:
                record.extend(['' for i in range(numcols-len(record))])
            try:
                unitcodes.remove(record[0])
            except ValueError: # code not in list
                logwrite('#unitcode in mapping file is not in subject basedata: '+record[0])
            outfile.write(','.join(record)+'\n')
        for c in unitcodes: # write out any remaining new codes not in existing file
            record = ['' for i in range(numcols)]
            record[0] = c
            outfile.write(','.join(record)+'\n')
        outfile.close()
        self.mapSubjects()

    def getMappingFileName(self):
        return self.app.getFullPath('OUTPATH')+'subjectmapping.csv'

    def getSIMSCodesfromPredictions(self):
        # braces to make a set and lose duplicates
        return {code for s in self.app.getStudentManager() for code, grade in s.getPredictions().items()}

    def getSubjectbyUnitCode(self, code):
        for s in self.getSubjects():
            if s.getUnitCode() == code:
                return s
        return None

    def getSubjectbySIMSName(self, simsname):
        for s in self.getSubjects():
            if s.getSIMSName() == simsname:
                return s
        return None

    def addSubjectfromBasedata(self, bdsubject):
        self.subjects.append(bdsubject)

#########################################################################################################
#
#   CLASS GUI Manager
#
# #########################################################################################################

class GUIManager:

    BROWSE_HEADINGS         = [ ['DOB', 'UCAS ID', 'Cycle', 'Postcode', 'UPN', 'UCI', 'ULN', 'ExamNo', 'New?', 'Y13?'],
                                ['Choice', 'Uni', 'Course', 'Status', 'Offer', 'Interview?'],
                                ['Code', 'Name', 'SIMS Code', 'Result', 'Prediction'] ]
    BROWSE_WIDTHS           = [ [1, 2], [1, 2, 5, 1, 2, 2], [3, 6, 4, 4, 4] ]
    BROWSE_JUSTIFYS         = [ [0, 0], [0, 0, 1, 0, 0, 0], [0, 1, 0, 0, 0] ]
    SEARCH_TABLE_SETTINGS   = [ ['Name', 'Course', 'Possible', 'Uni Code', 'Offer Grades', 'Predicted', 'Results'],
                                [6, 6, 2, 4, 3, 3, 3],
                                [1, 1, 0, 0, 0, 0, 0] ]

    def __init__(self, guimanager, studentmanager, subjectmanager):
        self.chosenstudent = 0
        self.guimanager = guimanager
        self.studentmanager = studentmanager
        self.subjectmanager = subjectmanager
        self.chosendate = self.studentmanager.getCurrentDate()

    def getFormattedDate(self):
        return self.formatDate(self.getDate())

    def formatDate(self, thedate):
        try:
            return thedate[:2]+'/'+thedate[2:4]+'/'+thedate[-4:]
        except:
            return 'None'

    def getDate(self):
        return self.chosendate

    def getStudent(self):
        return self.studentmanager.getStudentbyPosition(self.chosenstudent)

    def getStudentIndex(self):  # may only be used in browse bottom - remove later - also search - probs keep
        return self.chosenstudent

    def setStudentIndex(self, n):  # may only be used in search bottom - remove later?
        self.chosenstudent = n

    def incrementStudent(self):
        if self.chosenstudent < self.studentmanager.getNumStudents()-1:
            self.chosenstudent += 1
        return self.getStudent().getName()

    def decrementStudent(self):
        if self.chosenstudent > 0:
            self.chosenstudent -= 1
        return self.getStudent().getName()

    def resetDate(self):
        self.chosendate = self.studentmanager.getCurrentDate()

    def decrementDate(self):
        datelist = self.studentmanager.getAllDatesSeen()
        pos = datelist.index(self.chosendate)
        if pos < len(datelist)-1:
            pos += 1
            self.chosendate = datelist[pos]
        return self.chosendate

    def incrementDate(self):
        datelist = self.studentmanager.getAllDatesSeen()
        pos = datelist.index(self.chosendate)
        if pos > 0:     # low numbers are newer dates hence incrementing
            pos -= 1
            self.chosendate = datelist[pos]
        return self.chosendate

    def getPersonalData(self):
        student = self.getStudent()
        if student:
            d = []
            d.append(student.getDOBstring('%d/%m/%Y'))
            d.append(student.getUcasID())
            d.append(student.getCycle())
            d.append(student.getPCode())
            d.append(student.getUPN())
            d.append(student.getUCI())
            d.append(student.getULN())
            d.append(student.getExamNo())
            d.append('Yes' if student.isNew() else 'No')
            d.append('Yes' if student.isCurrentY13() else 'No')
        else:
            d = ['' for i in range(len(GUIManager.BROWSE_HEADINGS[0]))]
        return d

    def getChoiceData(self):
        student = self.getStudent()
        thedate = self.getDate()
        d = []
        if student is not None and thedate is not None:
            choices = student.getChoices(thedate)
            for i, c in enumerate(choices):
                d1 = []
                d1.append(c.getID())
                d1.append(c.getUni())
                d1.append(c.getCrsText())   # max field length [0:40] removed
                d1.append(c.getFullOutcome())
                d1.append(c.getOfferGrades(astar=True))
                wasinv = student.getInterviewDate(c.getID())
                if wasinv is not None:
                    d1.append(self.formatDate(wasinv))
                else:
                    d1.append('No')
                d.append(d1)
        return d

    def getResultData(self):
        student = self.getStudent()
        d = []
        if student is not None:
            usingresults = True
            collection = student.getResults()
            if len(collection) == 0:
                usingresults = False
                collection = student.getPredictions()   # if no results yet, go off predictions
            resultswithnomapping = False
            for k, v in collection.items():
                d1 = []
                if usingresults:
                    unit = k
                    result = v
                    unit = result.getUnitCode()
                    subject = self.subjectmanager.getSubjectbyUnitCode(unit)
                    if subject:
                        name = subject.getName()
                        simsname = subject.getSIMSName()
                        if simsname == '':
                            resultswithnomapping = True
                    else:               # basedata not imported
                        name = ''
                        simsname = ''
                    resultgrade = result.getGrade()
                else:
                    simsname = k
                    grade = v       # not needed as getPredbySIMS gets this value
                    unit = ''
                    name = ''
                    subject = self.subjectmanager.getSubjectbySIMSName(simsname)
                    if subject:     # user has done the mapping
                        unit = subject.getUnitCode()
                        name = subject.getName()
                    resultgrade = ''
                d1.append(unit)
                d1.append(name)    # removed [:15]
                d1.append(simsname)
                d1.append(resultgrade)
                d1.append(student.getPredictionbySIMSName(simsname))
                d.append(d1)
            if resultswithnomapping:    # show the unmapped sims codes and predictions Ma Ph Ch  AAA
                d1 = ['' for i in  range(len(GUIManager.BROWSE_HEADINGS[2]))]
                d1[2] = student.getPredictedSubjectCodeString()
                d1[3] = student.getPredictedGradeString()
                d.append(d1)
        return d

    def search(self, target, sortcolumn):
        dataset = []
        # Find list of student objects with matching personal or choice details
        matchlist = []
        thedate = self.getDate()
        for i, s in enumerate(self.studentmanager):
            # search in student data
            data = [ s.getName(),
                     s.getName(),
                     s.getUPN(),
                     s.getUcasID(),
                     s.getCycle(),
                     s.getDOBstring('%d%m%Y'),
                     s.getDOBstring('%d/%m/%Y'),
                     s.getDOBstring('%d-%m-%Y'),
                     s.getPCode() ]
            datastring = '#'.join(map(lambda x:x.lower(), data))
            if target in datastring:
                matchlist.append(i)
                continue
            # search in choices
            choices = s.getChoices(thedate)
            if choices:
                for c in choices:
                    data = [ c.getUni(),
                             c.getCrs(),
                             c.getCrsText(),
                             c.getOfferGrades(astar=True),
                             c.getFullOutcome() ]
                    datastring = '#'.join(map(lambda x:x.lower(), data))
                    if target in datastring and i not in matchlist:
                        matchlist.append(i)
            # search in results
            results = s.getResults()
            if results:
                for code, r in results.items():
                    subj = self.subjectmanager.getSubjectbyUnitCode(r.getUnitCode())
                    data = [ code,
                             subj.getName(),
                             subj.getSIMSName() ]
                    datastring = '#'.join(map(lambda x:x.lower(), data))
                    if target in datastring and i not in matchlist:
                        matchlist.append(i)
        # Now collect data to be displayed for those matching students
        if len(matchlist) != 0:   # otherwise "exit search" (in cleartable) gets overwritten
            for studentnumber in matchlist:
                s = self.studentmanager.getStudentbyPosition(studentnumber)
                studentdata = [studentnumber, s.getName(), '', '', '', '', '', '']
                if thedate is not None:
                    # get firm or, if not, choice #1
                    f = s.getFirm(thedate)
                    note = ' (CF)'
                    if not f:
                        # no firm so use first choice, if student has any choices at all
                        f = s.getChoices(thedate)
                        if f: f = f[0]
                        note = ' (#1)'
                    if f:
                        studentdata[2] = f.getCrsText()
                        studentdata[4] = f.getUni()+note
                        studentdata[5] = f.getOfferGrades(astar=True)
                    studentdata[3] = s.getPossibleOffers(thedate)
                    studentdata[6] = s.getPredictedGradeString()
                    studentdata[7] = s.getResultsAsOffer().getGrades(astar=True)
                dataset.append(studentdata)
            # sort according to selected column
            dataset.sort(key=lambda x:x[sortcolumn+1])
        return dataset

#########################################################################################################
#
#
#
#  CLASS TAURUS APPLICATION
#
#
#
#########################################################################################################

class TaurusApp():

    # ini file names
    INIFILE                  = 'TAURUS.ini'
    LICENCE                  = 'LICENCE.ini'
    # file ID stamp
    DATAFILETOKEN            = '#!TAURUSDATA'
    BASEFILETOKEN            = '#!TAURUSBASE'
    # licence hash salt
    LICSALT                  = '1234567890'
    # constants
    OPENFAIL                 = '#'      # anything that isn't a file handle

    def __init__(self):

        global logwrite
        logwrite = self.startupLogger

        # App-global data objects
        self.studentmanager = StudentManager(self)
        self.subjectmanager = SubjectManager(self)
        self.config = {}        # config parameters loaded below

        # Set-up and load data structures
        self.setConfigParameters()
        self.studentmanager.loadStudents()
        self.subjectmanager.loadSubjects()
        self.subjectmanager.mapSubjects()

        # Create GUI instance
        self.guimanager = None     # gui helper object
        self.gui = taurusGUI.TaurusGUI(self)

    def startupLogger(self, message):
        print(self.preprocesswarning(message))

    ##################################################################################
    #
    #   Startup and shutdown methods
    #
    ##################################################################################

    def run(self):
        try:
            self.gui.start()
        except Exception as e:
            logwrite('exception while initialising GUI')
            logwrite(str(e))
            self.gui.end()

    def quitApp(self):
        logwrite('saving data and closing application')
        try:
            if self.getStudentManager().getNumStudents() != 0:
                self.getStudentManager().saveStudents()
            if self.getSubjectManager().getNumSubjects() != 0:
                self.getSubjectManager().saveSubjects()
                self.getSubjectManager().updateSubjectMapping()
            self.gui.end()
        except Exception as e:
            logwrite('exception while closing app: ' + str(e))
            self.crashApp()

    def crashApp(self):
        logwrite('closing application due to error')
        try:
            self.gui.end()
        except AttributeError:  # die before gui created
            exit()

    def setConfigParameters(self):
        f = self.trytoopen(os.path.join('.',TaurusApp.INIFILE),
                           'error: configuration file not found - exiting')

        if f != TaurusApp.OPENFAIL:
            try:
                line = 'start of file'
                with f:
                    for line in f:
                        key, value = line.strip().split('=')
                        self.config[key] = value
            except ValueError:
                logwrite('warning: error in config file format - defaults used after: ' + line)
        else:
            self.crashApp()

        # Defaults to current year up to 2nd week in Sept
        # Should be done by then so default moves to the forthcoming year i.e. new academic year
        currentYear = datetime.datetime.strftime(datetime.datetime.now()+datetime.timedelta(weeks=16),' %Y')
        defaultSettings = [ ('ROOTPATH', '.'),
                            ('ASRPATH', '%U/downloads'),             # %U will be e.g. C:/users/jason/
                            ('PKLNAME', 'asrdata'),
                            ('PKLPATH', '%P/data'),
                            ('OUTNAME', 'report-'),
                            ('OUTPATH', '%U/taurus'),
                            ('APPYEAR', currentYear),
                            ('EXAMSIN', 'S:/SIMS/EXAMS/EXAMSIN'),
                            ('ESTABNO', '12345'),
                            ('CENTERN', '67890'),
                            ('LOGGING', '0')
                            ]

        # Set any parameters missing from the file using the above list
        # File MUST contain ESTABNO and CENTERN for license checking
        for k, v in defaultSettings:
            if k not in self.config:
                self.config[k] = v

        # If ASRNAME missing, use standard ucas format for asr file names
        if 'ASRNAME' not in self.config:
            self.config['ASRNAME'] = 'applicant_status_report_' + \
                                     self.config['APPYEAR'] + "_" + self.config['ESTABNO']
        # type conversions
        try:
            self.setConfig('LOGGING', int(self.getConfig('LOGGING')))
        except TypeError:
            self.setConfig('LOGGING', 0)    # default

        # debug
        for k, v in self.config.items():
            logwrite('#' + k + ": " + str(v))

        # License parameters
        f = self.trytoopen(os.path.join('.',TaurusApp.LICENCE),
                           'licence file not found or content error')
        if f == TaurusApp.OPENFAIL:
            self.crashApp()
        with f:
            licence = f.readline().strip()
            t = self.getLicenseToken()
        try:
            assert licence == t
        except AssertionError:
            logwrite('invalid license')
            self.crashApp()

    def getLicenseToken(self):
        return hashlib.sha256(bytes(self.getConfig('ESTABNO')+self.getConfig('CENTERN')+TaurusApp.LICSALT,'utf8')).hexdigest()

    def validateLicence(self, token, msg):
        if token != self.getLicenseToken():
            logwrite('license violation: ' + msg)
            self.crashApp()

    def getGUIManager(self):
        return self.guimanager

    def setGUIManager(self, guimanager_object):
        self.guimanager = guimanager_object

    ##################################################################################
    #
    #   UI interface
    #
    ##################################################################################

    def navigateTo(self, screen):
        if screen == 'browse' and not self.studentmanager.isLoaded():
            logwrite('cannot select browse with no data loaded')
            return
        else:
            if screen == 'browse':
                if self.getGUIManager() is None:
                    self.setGUIManager(GUIManager(self.gui.getBrowseLayout(), self.studentmanager, self.subjectmanager))
            self.gui.navigateTo(screen)

    def getBrowseHeadings(self):
        return GUIManager.BROWSE_HEADINGS

    def getBrowseWidths(self):
        return GUIManager.BROWSE_WIDTHS

    def getBrowseJustifys(self):
        return GUIManager.BROWSE_JUSTIFYS

    def getSearchTableSettings(self):
        return GUIManager.SEARCH_TABLE_SETTINGS

    def getGUIlabeldata(self):
        datestring = self.studentmanager.getCurrentDate()
        try:
            datestring = datestring[:2]+'/'+datestring[2:4]+'/'+datestring[-4:]
        except:
            datestring = 'None'
        return [ ('File Path', self.getConfig('ROOTPATH')),
                 ('Licensed to', self.getConfig('ESTABNO')),
                 ('Data Date', datestring) ]

    ##################################################################################
    #
    #   Get and set methods
    #
    ##################################################################################

    def getStudentManager(self):
        return self.studentmanager

    def getSubjectManager(self):
        return self.subjectmanager

    def getLogger(self):
        return self.gui.warning

    def verboseLogging(self):
        try:
            return self.getConfig('LOGGING')
        except:
            return 2        # LOGGING is set 0 by default so there's a problem; give all messages

    def preprocesswarning(self, message):
        # messages starting # only to be shown if VERBOSE logging > 1 - verbose mode
        # messages starting @ only to be shown if VERBOSE logging = 2 - DEBUG mode
        if message[0] in '#@':
            if self.verboseLogging() > '#@'.index(message[0]):
                caller = sys._getframe().f_back.f_back.f_code.co_name
                message = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ': ' +\
                          caller + ': ' + message.lstrip('#@')
            else:
                return ''
        else:
            message = message[0].upper() + message[1:]
        return message

    def getConfig(self,key):
        if key not in self.config:
            return 'NOTHING' ######## TEMP - remove when all retired keys sorted from GUI
        else:
            return self.config[key]

    def setConfig(self,key,value):
        self.config[key] = value

    ################################################################################
    #
    #  File handling methods
    #
    #################################################################################

    def trytoopen(self, filename, msg='', mode='r'):
        try:
            handle = open(filename,mode)
        except IOError:
            logwrite(msg.replace('%F', filename))
            return TaurusApp.OPENFAIL
        else:
            return handle

    def getFullPath(self, key):
        # Expand configured path name by replacing placeholders for user home / program root paths
        path = self.getConfig(key)
        if '%U' in path:
            return (os.path.expanduser(path.replace('%U','~'))).replace('\\','/')
        elif '%P' in path:
            return path.replace('%P',self.getFullPath('ROOTPATH'))
        else:
            return path

    def getReportFileName(self, report):
        thedate = self.studentmanager.getCurrentDate()
        if thedate is None:
            thedate = '0000' + self.getConfig('APPYEAR')
        return self.getFullPath('OUTPATH') + self.getConfig('OUTNAME') + \
               report + thedate + '.csv'

    def chooseFiletoOpen(self, options):
        filename = self.gui.fileOpenDialog(**options)
        lastslash = filename.rfind('/')
        if lastslash == -1:
            logwrite('#no backslash in filename OR user cancelled - ignoring')
            return None
        if not os.path.isfile(filename):
            logwrite("@chosen file isn't a file?")
            return None
        return filename

    def getASRimportfilenames(self):
        # Look in ASR path for files that contain data with dates not in the
        # current pickled dataset (which we loaded on App startup).
        # Returns list of filenames or None
        path = self.getFullPath('ASRPATH')
        asrfiles = [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f)) and f[-4:] == '.csv' and self.getConfig('ASRNAME') in f]
        if len(asrfiles)==0:
            logwrite('no ASR files found in ' + path)
            return None
        else:
            filelist = []
            for i, file in enumerate(asrfiles):
                f = self.trytoopen(path+file, 'skipping candidate ASR file %F - cannot open.')
                if f == TaurusApp.OPENFAIL:
                    continue
                with f:
                    thisDateString = f.readline()[42:52]
                    logwrite('found ASR file with date ' + thisDateString)
                    thisDateString.strip('/')
                    if thisDateString not in self.studentmanager.getAllDatesSeen():
                        filelist.append(path+file)
            if len(filelist) == 0:
                logwrite('already loaded any ASR files found in ' + path)
                return None
            else:
                return filelist

    def importASRdata(self):
        # Look for unloaded ASR files in default folder
        filelist = self.getASRimportfilenames()
        if filelist is None:
            return
        # load each one
        for defaultfile in filelist:
            logwrite('importing ASR data file: ' + defaultfile)
            # Finally got a valid file to open
            historic = False
            asr = ASRFile(self, defaultfile)
            if asr.getStatus() == TaurusApp.OPENFAIL:
                continue
            with asr:
                fileDate = asr.getFileDate()
                # convert date format to string for use below e.g. 23/07/2016 -> 23072016
                fileDateStr = fileDate.strftime('%d%m%Y')
                # Check if already imported
                if fileDateStr in self.studentmanager.getAllDatesSeen():
                    logwrite('#skipping already-imported file ' + defaultfile)
                else:
                    # Check if this is historic data
                    if len(self.studentmanager.getAllDatesSeen())>0:
                        if fileDate < datetime.datetime.strptime(self.studentmanager.getCurrentDate(),'%d%m%Y'):
                            historic = True
                    # License check
                    if asr.getEstabNo() != self.getConfig('ESTABNO'):
                        errstr = "#"+asr.getEstabNo()+"#"+self.getConfig('ESTABNO')+"#"
                        raise RuntimeError("Not licensed:"+errstr)
                    # Import data
                    # now add new file date to list of absorbed data dates
                    self.studentmanager.addDate(fileDateStr)
                    logwrite('#latest ASR date is now ' + self.studentmanager.getCurrentDate())
                    if len(self.studentmanager.getAllDatesSeen()) > 1:
                        logwrite('#previous ASR date is now ' + self.studentmanager.getPreviousDate())
                    for record in asr:
                        if asr.isStudent():
                            currentStudent = self.studentmanager.addStudent(historic,*asr.getStudentFields())
                        elif asr.isChoice():
                            # addchoice won't duplicate an existing choice if reimporting an asr
                            currentChoice = currentStudent.addChoice(fileDateStr,Choice(*asr.getChoiceFields()))
                            if currentChoice.isInterview():
                                if not currentStudent.addInterview(currentChoice.getID(), fileDateStr):
                                    logwrite('#interview date not updated for ' + currentStudent.getName() + ' at ' +
                                             currentChoice.getUni() + ' (was ' +
                                             currentStudent.getInterviewDate(currentChoice.getID()) +
                                             ' still INV at'  + fileDateStr + ')' )
                            # flag those choices which have changed status since previous available ASR
                            if len(self.studentmanager.getAllDatesSeen()) > 1:
                                # leave as 'new' if this is first ASR imported
                                if currentStudent.getChoices(self.studentmanager.getPreviousDate()) is not []:
                                    # student had choices in previous ASR so find them and see if updated this time
                                    previousChoices = currentStudent.getChoices(self.studentmanager.getPreviousDate())
                                    currentChoice.setChoiceUpdateStatus(previousChoices)
                        else:
                            # report unexpected line
                            logwrite('#unexpected line in file was ignored:\n' + ','.join(record))
                # end of IF already imported...
            # end WITH ... now add new file date to list of absorbed data dates
        # end FOR available files ... finished looping through available files
        # save and update gui
        self.studentmanager.saveStudents()
        self.gui.refreshData()

#######################################################################
#
#  Report Generation
#
#######################################################################

    def reportByStudent(self):
        # Create By Student report file: number of choices by decision type.
        if self.studentmanager.isLoaded():
            logwrite('starting report by student')
            report = ByStudentReport(self)
            outputfilename = self.getReportFileName('byStudent')
            report.run(outputfilename)
            logwrite('completed report by student to '+outputfilename)
        else:
            logwrite('no UCAS data loaded for report by student')

    def reportOffers(self, reportall):
        # Create Offers report file - all offers, or any changes to choices/outcomes from previous imported ASR.
        if reportall or len(self.studentmanager.getAllDatesSeen()) >= 2:
            typemsg = 'all' if reportall else 'updates'
            logwrite('starting offers report (' + typemsg + ')')
            report = OffersReport(self, reportall)
            outputfilename = self.getReportFileName('offers-' + typemsg)
            report.run(outputfilename)
            logwrite('completed offers report to ' + outputfilename)
        else:
            logwrite('need at least two ASR datasets loaded to report updates')

    def reportByUni(self):
        # Create By Uni report file - decisions for all students by university.
        if self.studentmanager.isLoaded():
            logwrite('starting report by uni')
            report = ByUniReport(self)
            outputfilename = self.getReportFileName('byUni')
            report.run(outputfilename)
            logwrite('completed report by uni to '+outputfilename)
        else:
            logwrite('no UCAS data loaded for report by uni')

    def reportBySubject(self):
        # Create By subject report file - results by subject.
        if self.studentmanager.isLoaded() and self.subjectmanager.getNumSubjects() > 0:
            logwrite('starting report by subject')
            report = BySubjectReport(self)
            outputfilename = self.getReportFileName('bySubject')
            report.run(outputfilename)
            logwrite('completed report by subject to '+outputfilename)
        else:
            logwrite('no UCAS or basedata loaded for report by subject')

    def reportAtRisk(self):
        # Create At Risk report file - students whose predictions are below firm offer etc.
        if self.studentmanager.isLoaded():
            logwrite('starting "at risk" report')
            report = AtRiskReport(self)
            outputfilename = self.getReportFileName('atRisk')
            report.run(outputfilename)
            logwrite('completed "at risk" report to '+outputfilename)
        else:
            logwrite('no UCAS data loaded for "at risk" report')

    def reportDestinations(self):
        # Create destinations report file - outcome for firm and insc by student
        if self.studentmanager.isLoaded():
            logwrite('starting destinations report')
            report = DestinationsReport(self)
            outputfilename = self.getReportFileName('destinations')
            report.run(outputfilename)
            logwrite('completed destinations report to '+outputfilename)
        else:
            logwrite('no UCAS data loaded for destinations report')

    ##########################################################################
    #
    # Data import and export routines - each uses a data source object
    #
    ##########################################################################

    def exportForSIMS(self):
        # Read exported SIMS marksheet (XML) and populate with UCAS data
        # Provided for centres who want to analyse UCAS data alongside / in SIMS
        # Check we have some data
        if not self.studentmanager.isLoaded():
            logwrite('no UCAS data loaded for export to SIMS')
            return
        logwrite('starting data collation for marksheet update')
        currentDate = self.studentmanager.getCurrentDate()
        # Create data table
        # assumed layout: UPN, Name, DOB, ExamNo, Unis, Courses, Offers, Outcomes, INVs
        mydata = []
        for s in self.studentmanager:
            unis = []
            courses = []
            offers = []
            outcomes = []
            wasinterviewed = []
            for c in s.getChoices(currentDate):
                unis.append(c.getUni())
                courses.append(c.getCrsText())
                offers.append(c.getOffer().getFullGrades())
                outcomes.append(c.getFullOutcome())
                wasinterviewed.append('Y' if s.getInterviewDate(c.getID()) is not None else '')
            unis.extend(['']*(5-len(unis)))
            courses.extend(['']*(5-len(courses)))
            offers.extend(['']*(5-len(offers)))
            outcomes.extend(['']*(5-len(outcomes)))
            wasinterviewed.extend(['']*(5-len(wasinterviewed)))
            # field positioning matters for SIMSXMLWriter
            rowdata = [s.getUPN(), s.getName(), s.getDOBstring('%d/%m/%Y'), s.getExamNo()]
            rowdata.extend(unis)
            rowdata.extend(courses)
            rowdata.extend(offers)
            rowdata.extend(outcomes)
            rowdata.extend(wasinterviewed)
            mydata.append(rowdata)
        # Read and update the XML marksheet
        opts = {'defaultextension': '.xml',
                'initialdir':self.config['ASRPATH'],
                'title':'Choose SIMS Marksheet file to import from'}
        infile = self.chooseFiletoOpen(opts)
        if not infile:
            logwrite('#cancelled by user')
            return
        opts['initialdir'] = self.config['OUTPATH']
        opts['title'] = 'Choose SIMS Marksheet file to export to'
        outfile = self.gui.fileSaveAsDialog(**opts)
        lastslash = outfile.rfind('/')
        if lastslash == -1:
            logwrite('#no backslash in outfilename OR user cancelled - ignoring')
            return
        g = self.trytoopen(outfile, 'cannot open output file - file open?', mode='wb')
        if g == TaurusApp.OPENFAIL:
            return
        logwrite('@begin SAX parse of input file')
        with g:
            SAX.parse(infile, SIMSXMLWriter(self, g, mydata))
        logwrite('success - marksheet updated and ready for import to SIMS')

    def importSIMSpredictions(self):
        # Read exported SIMS marksheet (XML) and extract predictions data
        # Relies on hidden UPN in SIMS marksheet export and configured column head for subject
        studentmanager = self.getStudentManager()
        # Check we have some data
        if not studentmanager.isLoaded():
            logwrite('no UCAS data loaded to associate with any imported SIMS data')
            return
        # Read and update the XML marksheet
        opts = {'defaultextension': '.xml',
                'initialdir':self.config['ASRPATH'],
                'title':'Choose SIMS Marksheet file to import from'}
        infile = self.chooseFiletoOpen(opts)
        if not infile:
            return
        logwrite('starting to read imported file')
        SXReader = SIMSXMLReader()
        SAX.parse(infile, SXReader)
        headerline = True
        for row in SXReader:
            if headerline:
                if 'upn' in map(lambda x:x.lower(), row):   # marksheet always has Upn in first column
                    headerline = False      # now seen headers, rest is students
                    headings = list(map(lambda x:x.lower(), row)) # convert to lower
                    # extract subject names using format given in app config
                    preamble, postamble = self.getConfig('PREDICT').lower().split('%')
                    for i, heading in enumerate(headings):
                        if heading.startswith(preamble) and heading.endswith(postamble):
                            headings[i] = heading[len(preamble):heading.index(postamble)].capitalize()
                        elif heading != 'upn' and 'name' not in heading.lower():
                            # keep upn and names but blank the rest
                            headings[i] = ''
            else:  # student row
                upn = None
                grades = {}
                for i, grade in enumerate(row):
                    if headings[i] == 'upn':
                        upn = grade
                    elif 'name' in headings[i].lower():     # implement find by name? need UCI anyway for results
                        name = grade
                    elif headings[i] != '':
                        if grade:                                   # ignore blanks
                            if grade.upper() in 'A*BCDE':
                                grades[headings[i]] = grade.upper() # can't add straight onto student as ...
                            else:
                                upn = row[headings.index('upn')]
                                logwrite('Ignored grade ' + grade + ', heading ' + headings[i] +
                                         ' for ' + str(studentmanager.getStudentbyUPN(upn)))
                me = self.studentmanager.getStudentbyUPN(upn)       #  ... UPN might not have been seen yet
                if not me:
                    logwrite('UPN '+str(upn)+' not matched: student not in UCAS or report not loaded?')
                    logwrite('Row begins ... '+','.join(row[0:5]))
                else:
                    for subject, grade in grades.items():
                        oldgrade = me.addPrediction(subject, grade)
                        if oldgrade is None:
                            logwrite('added ' + subject + ' ' + grade + ' to ' + me.getName())
                        elif oldgrade != False:
                            logwrite("Existing prediction was updated for " + me.getName()+" for "+subject+ \
                                          " from "+oldgrade+" to "+grade)
                        else:
                            pass    # same grade was imported again
        logwrite('success - predictions imported')
        logwrite('use Excel to update subject mappings file')
        self.getStudentManager().saveStudents()
        self.getSubjectManager().updateSubjectMapping()
        self.gui.refreshData()


    def importBasedata(self):
        subjects = self.getSubjectManager()
        logwrite('starting basedata import')
        basedata = BasedataDS(self)
        for bdsubject in basedata:
            if bdsubject and bdsubject.getQualLevel() == JCQ.SUBJECT_ALEVEL:
                subjects.addSubjectfromBasedata(bdsubject)
        logwrite('success - subjects added')
        logwrite('use Excel to update subject mappings file')
        subjects.saveSubjects()
        subjects.updateSubjectMapping()
        self.gui.refreshData()

    def importResults(self):
        subjects = self.getSubjectManager()
        studentmanager = self.getStudentManager()
        if len(subjects.getSubjects()) == 0:
            logwrite('no subject basedata loaded')
            return
        logwrite('starting results import')
        results = ResultDS(self)
        for result in results:
            if result:
                # only import full A level results
                subj = subjects.getSubjectbyUnitCode(result.getUnitCode())
                if not subj:
                    logwrite('ignoring result for subject not in basedata: ' + result.getUnitCode() +
                             ' for ' + str(studentmanager.getStudentfromResult(result)))
                else:
                    if subj.getQualLevel() == JCQ.SUBJECT_ALEVEL:
                        s = studentmanager.getStudentfromResult(result)
                        if s:
                            if not s.addResult(result):
                                logwrite('duplicate grade in results file: ' \
                                    + str(s) + ' already has record ' \
                                    + str(s.getResultbyUnit(result.getUnitCode())) \
                                    + str(subj) \
                                    + ': skipped adding new record')
                        else:
                            logwrite('student in results file is not found in loaded data: ' + result.getUCI())
                    else:
                        logwrite('non A level result ignored: ' + subj.getQualLevel() + ' ' + subj.getName())
        logwrite('success - results imported')
        logwrite('select "Browse" or create destinations report to analyse results')
        studentmanager.saveStudents()
        self.gui.refreshData()

    def importFromSIMS(self):
        # Get student details (UPN etc) not available from UCAS from SIMSExtractDS
        # Check we have some UCAS data to cross-reference
        studentmanager = self.getStudentManager()
        if not studentmanager.isLoaded():
            logwrite('no UCAS data loaded to associate with any imported SIMS data')
            return
        logwrite('starting student details import')
        SIMSreportData = SIMSExtractDS(self)  # pass app object
        for student in SIMSreportData:
            if student is not None:
                logwrite('#student record was updated for ' + student.getName())
        logwrite('success - student details added')
        studentmanager.saveStudents()
        self.gui.refreshData()


#########################################################################################################
#
#
#  CLASSES for REPORTING
#
#
#########################################################################################################

class StudentReport():  # base class provides common init and save methods

    def __init__(self, app, *args, **kwargs):
        self.app = app
        self.records = []
        self.studentmanager = app.getStudentManager()
        self.subjectmanager = app.getSubjectManager()

    def run(self, outputfile):
        self.format()           # overridden by subclass depending on report type
        self.save(outputfile)

    def save(self, filename):  # must not be called on the base class as no 'headings' member
        f = self.app.trytoopen(filename, 'report file %F may be open or protected: data not written', mode='w')
        if f != TaurusApp.OPENFAIL:
            with f:
                f.write(",".join(self.headings)+'\n')
                for record in self.records:
                    line = ''
                    for heading in self.headings:
                        line += str(record[heading])+','
                    f.write(line[:-1]+'\n')  # remove last comma

    def getItems(self, choice): # used by both Destinations and AtRisk reports
        if choice is None:
            return ['None', 'None', 'N/A']
        else:
            items = [choice.getCrsText(), choice.getUni()]
            if choice.getOutcome() == Outcome.C:
                items.append(choice.getOfferGrades(astar=True))
            elif choice.getOutcome() == Outcome.U:
                items.append("U")
            else:
                logwrite('@offer choice is neither C nor U: internal error (' + ','.join(items) + ')')
                items.append('Error: not offer choice?')
            return items

    def prettifyOutcome(self, outcome):
        if outcome == Outcome.C:
            return 'Conditional Offer'
        elif outcome == Outcome.U:
            return 'Unconditional Offer'
        elif outcome == Outcome.REJ:
            return 'Rejected'
        elif outcome == Outcome.INV:
            return 'Invited to Interview'
        elif outcome == Outcome.REF:
            return 'Referred'
        elif outcome == Outcome.W:
            return 'Withdrawn'
        else:
            logwrite('@invalid outcome ' + outcome)
            return None

    def prettifyStatus(self, updated):
        result = ''
        if updated == Update.UPD8_NEW:
            result = 'NEW CHOICE'
        else:
            if updated & Update.UPD8_COURSE:
                result = 'NEW COURSE'
            if updated & Update.UPD8_OUTCOME:
                result += ' & ' + 'NEW OUTCOME'
            if len(result)>1 and result[1] == '&':
                result = result[3:] # strip ampersand if newoutcome was the only change
        if result == '' and updated != Update.UPD8_SAME:  # either UPD8_UNDEFINED or something else?
            logwrite('@invalid update value ' + str(updated))
            return None
        return result

class OffersReport(StudentReport):

    def __init__(self, *args, **kwargs):
        self.reportall = args[1]    # arg 0 is the app object
        # not reportall means just show updates since the previous ASR
        super().__init__(*args, **kwargs)
        self.headings = ['Name', 'Course', 'Uni', 'Outcome', 'Offer Grades']
        if not self.reportall:
            self.headings.append('Status')

    def format(self):
        for student in self.studentmanager:
            record = {}
            if student.isNew(): # show new applicants whether or not reporting all offers
                for heading in self.headings:
                    record[heading] = ''
                record[self.headings[0]] = student.getName()
                index = 4 if self.reportall else 5
                record[self.headings[index]] = 'NEW APPLICANT'
                self.records.append(record)
            else:
                choices = student.getChoices(self.studentmanager.getCurrentDate())
                if choices:
                    for choice in choices: # report all offers or just updated choices
                        if (not self.reportall and choice.hasUpdated()) or (self.reportall and choice.isOffer()):
                            record = {	self.headings[0]: student.getName(),
                                        self.headings[1]: choice.getCrsText(),
                                        self.headings[2]: choice.getUni(),
                                        self.headings[3]: self.prettifyOutcome(choice.getOutcome()),
                                        self.headings[4]: choice.getOfferGrades(astar=True)  }
                            if not self.reportall:
                                record[self.headings[5]] = self.prettifyStatus(choice.getUpdated())
                            self.records.append(record)

class BySubjectReport(StudentReport):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.headings = ['Subject Name', 'SIMS Code', 'Unit', 'A*', 'A', 'B', 'C',
                         'D', 'E', 'U', 'Other', 'UnderPrediction', 'OverPrediction']

    def format(self):
        ConvertGrade = lambda x:'U E D C B A A*'.index(x)//2
        records = {}    # empty dict
        for subject in self.subjectmanager.getSubjects():
            if subject.getSIMSName():
                n = subject.getName()
                if ',' in n:
                    n = '"'+n+'"'
                record = { self.headings[0]: n,
                           self.headings[1]: subject.getSIMSName(),
                           self.headings[2]: subject.getUnitCode(),
                           self.headings[3]: 0,
                           self.headings[4]: 0,
                           self.headings[5]: 0,
                           self.headings[6]: 0,
                           self.headings[7]: 0,
                           self.headings[8]: 0,
                           self.headings[9]: 0,
                           self.headings[10]: 0,
                           self.headings[11]: 0,
                           self.headings[12]: 0,
                           }
                records[subject.getUnitCode()] = record
        for student in self.studentmanager:
            studentresults = student.getResults()
            studentpredictions = student.getPredictions()
            for code, result in studentresults.items():
                try:
                    gindex = ConvertGrade(result.getGrade())
                except ValueError:
                    gindex = 10         # other
                if code in records:     # only count results for which there are predictions
                    records[code][self.headings[9-gindex]] += 1
                    simscode = self.subjectmanager.getSubjectbyUnitCode(code).getSIMSName()
                    try:
                        predictedgrade = studentpredictions[simscode]
                    except KeyError:    # no prediction for this (retake?)
                        predictedgrade = studentresults[code].getGrade() # just count it as 0
                    try:
                        diff = ConvertGrade(studentresults[code].getGrade()) - ConvertGrade(predictedgrade)
                    except ValueError:
                        logwrite('grade '+studentresults[code].getGrade()+' or '+predictedgrade+' not a grade!')
                    if diff > 0:
                        records[code][self.headings[11]] += diff
                    else:
                        records[code][self.headings[12]] += -diff
                else:
                    logwrite('code '+str(code)+' not in '+str(studentresults)+' for '+student.getName())
        self.records = []
        for r in records:
            self.records.append(records[r])

class ByStudentReport(StudentReport):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.headings = ['Name','U','C','INV','REF','REJ','W','Other','Choice Status']

    def format(self):
        currentDate = self.studentmanager.getCurrentDate()
        for student in self.studentmanager:
            record = {	self.headings[0]: student.getName(),
                          self.headings[1]: student.getUnconditionals(currentDate),
                          self.headings[2]: student.getConditionals(currentDate),
                          self.headings[3]: student.getInterviews(currentDate),
                          self.headings[4]: student.getReferrals(currentDate),
                          self.headings[5]: student.getRejections(currentDate),
                          self.headings[6]: student.getWithdrawals(currentDate)  }
            totalsofar = sum((record[self.headings[i]] for i in range(1,7)))
            record[self.headings[7]] = student.getTotalChoices(currentDate) - totalsofar
            status = ''
            if student.getFirm(currentDate) is not None:
                status = 'CHOICES MADE (F=' + student.getFirm(currentDate).getOfferGrades(astar=True)
                if student.getInsc(currentDate) is None:
                    if student.getFirm(currentDate).getOutcome() != Outcome.U:
                        logwrite('warning: CF with no insurance for ' + student.getName())
                else:
                    status += ' I=' + student.getInsc(currentDate).getOfferGrades(astar=True)
                    if student.acceptanceAnomaly(currentDate):
                        logwrite('insurance offer higher than firm for ' + student.getName()
                                  + ': firm grades = ' + student.getFirm(currentDate).getOffer().getFullGrades()
                                  + ', insc grades = ' + student.getInsc(currentDate).getOffer().getFullGrades()  )
                status += ')'
            else:
                if student.getDecisions(currentDate) == student.getTotalChoices(currentDate):
                    if student.getTotalOffers(currentDate) == 0:
                        status = 'IN CLEARING'
                    elif student.getOpenOffers(currentDate) == 0:
                        status = 'DECLINED ALL'
                    else:
                        status = 'READY TO CHOOSE'
            record[self.headings[8]] = status
            self.records.append(record)

class ByUniReport(StudentReport):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Get a list of universities for headings
        # ... the {} makes it a set which removes duplicates
        current = self.studentmanager.getCurrentDate()
        self.universities = [UniRecord(u) for u in sorted({c.getUni() for s in self.studentmanager
                                                           for c in s.getChoices(current)})]
        for u in self.universities:
            u.setUnconditionals(self.countOutcomes(u, current, Outcome.U))
            u.setConditionals(self.countOutcomes(u, current, Outcome.C))
            u.setRejections(self.countOutcomes(u, current, Outcome.REJ))

        for s in self.studentmanager:
            for c in s.getChoices(current):
                if c.getOutcome() == Outcome.C:
                    u = self.universities[self.universities.index(c.getUni())]
                    u.addOffer(c.getOfferGrades())

        self.headings = list(map(lambda x: x.getName(), self.universities))
        self.headings.insert(0, 'Total')
        self.headings.insert(0, ' ')

    def format(self):
        self.prepareTotalsByUni()
        self.prepareBreakdownByOfferConditions()
        
    def countOutcomes(self, u, thisdate, thisoutcome):
        return sum([1 for s in self.studentmanager for c in s.getChoices(thisdate)
                         if c.getUni() == u.getName() and c.getOutcome() == thisoutcome])

    def prepareTotalsByUni(self):
        funcs = [lambda x:x.getUnconditionals(),lambda x:x.getConditionals(),lambda x:x.getRejections(),lambda x:0 if x.getTotalOutcomes()==0 else (x.getConditionals()+x.getUnconditionals())/x.getTotalOutcomes()]
        for i in range(4):
            record = {}  # empty record for this report row
            # set first field to row heading
            record[self.headings[0]] = "UCR%"[i]
            # get value for total column (heading 1)
            func = funcs[i]
            if i==3:  # percentage calculation
                v = (sum(map(funcs[0],self.universities))
                     + sum(map(funcs[1],self.universities))) \
                    / sum(map(lambda x:x.getTotalOutcomes(),self.universities))
                record[self.headings[1]] = str(int(v*100))
            else:  # just counting
                v = sum(map(func,self.universities))
                record[self.headings[1]] = str(v)
            # get value for each uni column
            for u in self.universities:
                record[u.getName()] = str(int(func(u)*(1 if i!=3 else 100)))
            self.records.append(record)

    def prepareBreakdownByOfferConditions(self):
        # set of all offer conditions
        offers = {(oc.getGrades(astar=True),oc.getGradeValue()) for u in self.universities
                                                        for oc in u.getAllOffers()}
        # breakdown of offers by uni
        for grades, value in sorted(list(offers), key=lambda x:x[1], reverse=True):
            record = {}
            # add row heading
            if grades == '':
                record[self.headings[0]] = 'Other'
            else:
                record[self.headings[0]] = grades
            # add field for each uni
            record[self.headings[1]] = ''  # blank underneath Total column
            for u in self.universities:
                count = u.getNumOffers(grades)
                record[u.getName()] = '' if count == 0 else str(count)
            self.records.append(record)

class DestinationsReport(StudentReport):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.headings = ['Name','Outcome','Current Y13?','UCASID','CycleNo',
                         'FirmCourse','FirmUni','FirmGrades','MetFirm?',
                         'InscCourse','InscUni','InscGrades','MetInsc?',
                         'ExamNo','ResultGrades','PredGrades','MetPred?', 'UPN']

    def format(self):
        currentDate = self.studentmanager.getCurrentDate()
        for student in self.studentmanager:
            record = {  self.headings[0]: student.getName(),
                          self.headings[1]: '',
                          self.headings[2]: student.isCurrentY13(),
                          self.headings[3]: student.getUcasID(),
                          self.headings[4]: student.getCycle()  }

            firm = student.getFirm(currentDate)
            insc = student.getInsc(currentDate)
            if insc is None:  # can't set I on UCAS without F
                if firm is None:
                    logwrite('#no firm (or insc) for ' + student.getName())
                elif firm.getOutcome() != Outcome.U:
                    logwrite('#CF but no insc for ' + student.getName())
            firmitems = self.getItems(firm)
            inscitems = self.getItems(insc)
            for i in range(3):
                record[self.headings[5+i]] = firmitems[i]
                record[self.headings[9+i]] = inscitems[i]
            record[self.headings[13]] = student.getExamNo()
            resoff = student.getResultsAsOffer()
            record[self.headings[14]] = resoff.getGrades(astar=True)
            predoff = student.getPredictionsAsOffer()
            record[self.headings[15]] = predoff.getGrades(astar=True)
            record[self.headings[16]] = Offer.DESCRIPTIONS[resoff.gradeCompare(predoff, False)]
            record[self.headings[17]] = student.getUPN()
            if firm:
                record[self.headings[8]] = Offer.DESCRIPTIONS[resoff.gradeCompare(firm.getOffer())]
            else:
                record[self.headings[8]] = Offer.DESCRIPTIONS[Offer.NOOFFER]
            if insc:
                record[self.headings[12]] = Offer.DESCRIPTIONS[resoff.gradeCompare(insc.getOffer())]
            else:
                record[self.headings[12]] = Offer.DESCRIPTIONS[Offer.NOOFFER]
            if firm is None:
                record[self.headings[1]] = 'No offers'
            elif resoff.gradeCompare(firm.getOffer()) in Offer.MET:
                record[self.headings[1]] = 'Firm'
            elif insc and resoff.gradeCompare(insc.getOffer()) in Offer.MET:
                record[self.headings[1]] = 'Insc'
            elif resoff.gradeCompare(firm.getOffer()) in Offer.UNMET \
                    and insc and resoff.gradeCompare(insc.getOffer()) in Offer.UNMET:
                record[self.headings[1]] = 'Unmet'
            else:
                record[self.headings[1]] = 'CHECK'
            self.records.append(record)

class AtRiskReport(StudentReport):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.headings = ['Name','FirmCourse','FirmUni','FirmGrades',
                         'InscCourse','InscUni','InscGrades','Predictions','Risk']

    def format(self):
        currentDate = self.studentmanager.getCurrentDate()
        first = True
        usingresults = False
        for student in self.studentmanager:
            record = {  self.headings[0]: student.getName() }
            firm = student.getFirm(currentDate)
            insc = student.getInsc(currentDate)
            firmitems = self.getItems(firm)
            inscitems = self.getItems(insc)
            fv = 0
            iv = 0
            if firm:
                fv = firm.getOffer().getGradeValue()
                if firm.getOutcome() == Outcome.U:
                    fv = 0
            if insc:
                iv = insc.getOffer().getGradeValue()
            n = 3 # number of predicted grades to consider as counting
            if firm: n = max(firm.getOffer().numGrades(), n)
            if insc: n = max(insc.getOffer().numGrades(), n)
            # report works off actual results if available, otherwise predictions
            try:
                pg = student.getResultsAsOffer().getGradeEquivalent(astar=True)
                pv = student.getResultsAsOffer().getGradeValue()
                if first:
                    logwrite('using actual results to determine risk status')
                    self.headings[self.headings.index('Predictions')] = 'Results'
                    usingresults = True
                    first = False
            except:
                pg = student.getPredictedGradeString(n)
                pv = Offer(pg).getGradeValue()
                if first:
                    logwrite('using predictions to determine risk status')
                    first = False
            # this report contains:
            # students with no firm
            # students with a firm offer > predictions
            # students with a firm = predictions and insc >= predictions
            # students with a firm = predictions and no insc
            # determine how at risk this student is
            risk = ''
            if firm is None: risk = 'HIGH'
            if fv > pv:
                risk = 'HIGH'
                if fv == 1000: risk += ' (special conditions)'
                if insc:
                    if iv == pv:
                        if not usingresults:
                            risk = 'MEDIUM'
                        elif insc.getOffer().getGrades() != pg:
                            risk = 'LOW'
                        elif usingresults: # the HIGH can become none as she's got her insc
                            risk = ''
                    elif iv < pv:
                        if not usingresults:
                            risk = 'LOW'
                        else:
                            risk = ''   # got insc
            if fv > 0 and fv == pv and not usingresults:
                if iv >= pv or insc is None:
                    risk = 'MEDIUM'
                    if iv == 1000: risk = 'MEDIUM (special conditions)'
                elif iv == pv - 1:
                    risk = 'LOW'
                # else FV=PV and IV at least 2 below PV so no risk
            # only copy at risk students to the report
            if risk:
                for i in range(3):
                    record[self.headings[1+i]] = firmitems[i]
                    record[self.headings[4+i]] = inscitems[i]
                record[self.headings[7]] = pg
                record[self.headings[8]] = risk
                self.records.append(record)

#########################################################################################################
#
#  CLASS SIMSXML Reader and Writer classes
#
#########################################################################################################

class SIMSXMLWriter(SAX.handler.ContentHandler):

    def __init__(self, app, outfile, student_data):
        self.app = app
        self.row = -1
        self.cell = -1
        self.chars = []
        self.students = student_data
        # columns in student data fixed in caller: XML sheet (SIMS empty export) must have standard layout
        self.keycolumns = [0,1,2,3]
        self.UPNcolumn, self.namecolumn, self.dobcolumn, self.examcolumn = self.keycolumns
        self.studentID = None
        self.hadDataTag = False
        self.gotExamNo = None
        self.output = SAXUTILS.XMLGenerator(outfile, 'utf-8')

    def startDocument(self):
        self.output.startDocument()

    def processingInstruction(self, target, data):
        self.output.processingInstruction(target, data)

    def endDocument(self):
        self.output.endDocument()

    def characters(self, content):
        self.chars.append(content)

    def startElement(self, name, attrs):
        if name=="Cell":
            self.hadDataTag = False
            self.cell += 1
            if 'ss:Index' in attrs:
                index = int(attrs['ss:Index'])
                self.cell = index
        elif name=="Row":
            self.cell = -1
            self.row += 1
        self.output.characters(''.join(self.chars))
        self.chars=[]
        self.output.startElement(name, attrs)

    def endElement(self, name):
        # end of row - clear 'memory' variables in case it was the last row
        if name == 'Row':
            self.hadDataTag = False
            self.studentID = None
        # collect data from between tags
        data = ''.join(self.chars)
        # student ID cells are: 0,1,2,3 = upn, names, dob and exam no
        # endelement will be called per Excel cell:
        # once if <cell> has no <data> inside
        # twice if both <cell> has <data> inside
        if self.cell <= 3:
            # in these cells try to identify which student this row is
            if self.studentID is None:
                self.identifyStudent(data)
            # now just output what's already there because
            # student ID won't be changed and any other rows mustn't be altered
            self.output.characters(data)
            self.chars = []
            self.output.endElement(name)
        elif self.studentID:
            # currently working on a student row
            if name == 'Cell' and not self.hadDataTag:
                # end of cell with no data inside: need to output the
                # passed-in data inside a data tag and close the cell also
                self.output.startElement('Data', attrs={'ss:Type':'String'})
                self.output.characters(self.studentID[self.cell])     #   write passed-in data
                self.output.endElement('Data')
                self.output.endElement('Cell')
                self.chars = []
                # will be starting a new cell so clear the flag
                self.hadDataTag = False
            else:
                # end of cell on student row but had data tag - pass on to just output
                if name == 'Cell' and self.hadDataTag:
                    pass
                else:   # otherwise note that we saw a data tag if we did
                    if name == 'Data':
                        self.hadDataTag = True
                        self.output.characters(self.studentID[self.cell])
                self.chars = []
                self.output.endElement(name)
        else:   # anything else just copy out
            self.output.characters(data)
            self.chars = []
            self.output.endElement(name)

    def identifyStudent(self, data):
        if self.cell == self.UPNcolumn:
            for s in self.students:
                if data == s[0]:        # UPN match
                    self.studentID = s
                    logwrite('@set student from upn')
                    break
        elif self.cell == self.namecolumn and len(data) != 0:
            for s in self.students:
                if data.lower() == s[1].lower()[:len(data)]:    # SIMS surname forename exact match
                    self.studentID = s
                    logwrite('@set student from name')
                    break
        elif self.cell == self.examcolumn:
            for s in self.students:
                if data == s[3]:        # exam number matches (probably ok) ...
                    self.gotExamNo = s
                    logwrite('@temp set student from exno')
                    break
        elif self.cell == self.dobcolumn:
            for s in self.students:     # ... and dob matches too - that'll do
                if data == s[2]:
                    if self.gotExamNo == s:
                        self.studentID = s
                    logwrite('@set student from examno/dob')
                    break
            self.gotExamNo = None

class SIMSXMLReader(SAX.handler.ContentHandler):
    def __init__(self):
        self.chars=[]
        self.cells=[]
        self.rows=[]
        self.ptr = -1

    def __next__(self):
        self.ptr += 1
        if self.ptr >= len(self.rows):
            raise StopIteration
        else:
            return self.rows[self.ptr]

    def __iter__(self):
        return self

    def characters(self, content):
        self.chars.append(content)

    def startElement(self, name, atts):
        if name=="Cell":
            self.chars=[]
        elif name=="Row":
            self.cells=[]

    def endElement(self, name):
        if name=="Cell":
            self.cells.append(''.join(self.chars))
        elif name=="Row":
            self.rows.append(self.cells)

#########################################################################################################
#
#
#
#  MAIN PROGRAM
#
#
#
#########################################################################################################

if __name__ == "__main__":
    logwrite = print            # function in global namespace used for logging
    app = TaurusApp()
    logwrite = app.getLogger()  # now redirect to gui-specific logger
    app.run()
