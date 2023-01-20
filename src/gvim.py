import api
import appModuleHandler
import tones
import comHelper
import controlTypes
import editableText
import review
import speech
import textInfos
import time
import ui
import comtypes.client

vim=None
def _vimEval(str):
	global vim
	return vim.Eval(str)

def _vimInt(str):
	return int(_vimEval(str))

def _line2offset(lnum):
	return _vimInt("line2byte(%s)"%lnum)

def _offset2line(offset):
	return _vimInt("byte2line(%s)"%offset)

oldLine=0
oldCol=0
lastCmdLine=""

def _getPos():
	pos=_vimEval("getpos('.')").splitlines()
	return (int(pos[1]),int(pos[2]))

def _getEOL():
	fileFormat=_vimEval("&fileformat")
	if fileFormat=="": fileFormat="dos"
	if fileFormat=="dos": eol="\r\n"
	elif fileFormat=="unix": eol="\n"
	else: eol="\r"
	return eol

class VimTextInfo(textInfos.offsets.OffsetsTextInfo):
	def _getStoryLength(self):
		result=_vimInt("line2byte(line('$')+1)")-1
		if result==-2: result=0
		return result

#	def _getStoryText(self):
#		return _vimEval("getline(1,'$')")

	def _getCaretOffset(self):
		return _vimInt("line2byte(line('.'))+col('.')-2")

	def _setCaretOffset(self, offset):
		offset+=1
		lnum=_offset2line(offset)
		lineStart=_line2offset(lnum)
		col=offset-lineStart+1
		_vimEval("cursor(%s,%s)"%(lnum,col))

	def _getLineOffsets(self, offset):
		offset+=1
		eol=_getEOL()
		lineNum=_offset2line(offset)
		lineStart=_line2offset(lineNum)
		lineEnd=_line2offset(lineNum+1)
		if _vimInt("&wrap")==1:
			columns=_vimInt("winwidth(0)")
			textWidth=_vimInt("strwidth(getline(%s))"%(lineNum))
			if textWidth>columns:
				offset-=lineStart
				lineStart+=(offset/columns)*columns
				if lineStart+columns<=lineEnd: lineEnd=lineStart+columns
		return[lineStart-1,lineEnd-1]

	def _getLineNumFromOffset(self, offset):
		offset+=1
		return _offset2line(offset)

	def _getTextRange(self, start, end):
		storyLength=self._getStoryLength()
		if start>=end or start<0 or start>storyLength or end<0 or end>storyLength: return ""
		start+=1
		end+=1
		lastLine=_vimInt("line('$')")
		eol=_getEOL()
		startLine=_offset2line(start)
		if startLine==-1: startLine=lastLine
		endLine=_offset2line(end)
		if endLine==-1: endLine=lastLine
		startOffset=_line2offset(startLine)
		lines=_vimEval("getline(%s,%s)"%(startLine,endLine)).splitlines()
		if lines==[]: return ""
		text=eol.join(lines)+eol
		endOffset=startOffset+len(text)
		i=start-startOffset
		j=endOffset-end
		if j==0: j=-len(text)
		return text[i:-j]

class VimCmdLineTextInfo(textInfos.offsets.OffsetsTextInfo):
	def _getStoryText(self):
		return _vimEval("getcmdtype()")+_vimEval("getcmdline()")

	def _getStoryLength(self):
		return len(self._getStoryText())

	def _getTextRange(self, start, end):
		text=self._getStoryText()
		return text[start:end]

	def _getCaretOffset(self):
		return _vimInt("getcmdpos()")

	def _setCaretOffset(self, offset):
		offset+=1
		_vimEval("setcmdpos(%s)"%offset)

class Vim(editableText.EditableTextWithoutAutoSelectDetection):
	def _get_TextInfo(self):
		if _vimEval("getcmdtype()")=="":
			return VimTextInfo
		else:
			return VimCmdLineTextInfo

	def event_typedCharacter(self, ch):
		global lastCmdLine
		if _vimEval("getcmdtype()")!="":
			lastCmdLine=_vimEval("getcmdline()")
		if _vimEval("mode()")=="i" or self.TextInfo is VimCmdLineTextInfo:
			speech.speakTypedCharacters(ch)
		else:
			#time.sleep(.05)
			global oldLine,oldCol
			line,col=_getPos()
			if line==oldLine and col==oldCol:
				speech.speakTypedCharacters(ch)
				return
			if line==oldLine:
				if abs(col-oldCol)==1: unit=textInfos.UNIT_CHARACTER
				else: unit=textInfos.UNIT_WORD
			else:
				unit=textInfos.UNIT_LINE 
			#self._caretScriptPostMovedHelper(unit,None,None)
			info = self.makeTextInfo(textInfos.POSITION_CARET)
			review.handleCaretMove(info)
			info.expand(unit)
			#speech.speakTextInfo(info, unit, reason=controlTypes.REASON_CARET)
			speech.speakTextInfo(info, unit, reason=controlTypes.OutputReason.CARET)
			oldLine,oldCol=line,col

	def script_reportCompletion(self, gesture):
		global lastCmdLine
		gesture.send()
		time.sleep(.05)
		if _vimEval("getcmdtype()")!="":
			l=len(lastCmdLine)
			ui.message(_vimEval("getcmdline()")[l:])

	def script_reportStatusLine(self, gesture):
		lines=[]
		windowWidth=_vimInt("winwidth(0)")
		self.redraw()
		text=self.displayText
		while len(text)>0:
			lines.append(text[:windowWidth])
			text=text[windowWidth:]
		if len(lines)==0: lines.append("")
		ui.message(lines[-1])
	__gestures={
		"kb:nvda+end":"reportStatusLine",
		"kb:tab":"reportCompletion",
	}

class AppModule(appModuleHandler.AppModule):
	def __init__(self, processID, appName=None):
		super(AppModule,self).__init__(processID,appName)
		global vim
		vim=comHelper.getActiveObject("Vim.Application",True)
		#vim=comtypes.client.GetActiveObject("Vim.Application")
		


	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		if obj.windowClassName==u'Vim':
			clsList.insert(0,Vim)
			#tones.beep(512, 100)

