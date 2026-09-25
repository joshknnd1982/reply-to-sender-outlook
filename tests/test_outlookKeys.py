"""Tests for Reply to Sender's keys and window handling as a global plugin (0.4).

python -m unittest discover -s tests

Up to 0.3, NVDA+shift+r and NVDA+shift+e were scripts of an Outlook app module,
which replaced Outlook Extended's (see test_outlookAddons). In 0.4 they belong to
the global plugin, which must keep them to classic Outlook: anywhere else NVDA
and other add-ons get the keys. The plugin is loaded here with NVDA 2026.2's own
scripting code, copied word for word below from baseObject.py,
globalPluginHandler.py, scriptHandler.py and inputCore.py; the rest of NVDA is
stubbed.
"""

import builtins
import importlib.util
import os
import sys
import textwrap
import types
import typing
import unittest
from unittest import mock

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLUGIN_DIR = os.path.join(ROOT, "addon", "globalPlugins", "replyToSenderOutlook")

# NVDA 2026.2 code, word for word.
SCRIPTABLE_TYPE = r'''
class ScriptableType(AutoPropertyType):
	"""A metaclass used for collecting and caching gestures on a ScriptableObject"""

	def __new__(cls, name: str, bases: tuple[type, ...], namespace: dict[str, Any], /, **kwargs: Any):
		newCls = super().__new__(cls, name, bases, namespace, **kwargs)
		gesturesDictName = "_%s__gestures" % newCls.__name__
		# #8463: To avoid name mangling conflicts, create a copy of the __gestures dictionary.
		try:
			gestures = getattr(newCls, gesturesDictName).copy()
		except AttributeError:
			# This class currently has no gestures dictionary,
			# because no custom __gestures dictionary has been defined.
			gestures = {}
		for name, script in namespace.items():
			if not name.startswith("script_"):
				continue
			scriptName = name[len("script_") :]
			if hasattr(script, "gestures"):
				for gesture in script.gestures:
					gestures[gesture] = scriptName
		if gestures:
			setattr(newCls, gesturesDictName, gestures)
		return newCls'''

SCRIPTABLE_OBJECT_INIT = r'''
def __init__(self):
	#: Maps input gestures to script functions.
	#: @type: dict
	self._gestureMap = {}
	# Bind gestures specified on the class.
	# This includes gestures specified on decorated scripts.
	# This does not include the gestures that are added when creating a DynamicNVDAObjectType.
	for cls in reversed(self.__class__.__mro__):
		try:
			self.bindGestures(getattr(cls, "_%s__gestures" % cls.__name__))
		except AttributeError:
			pass
		try:
			self.bindGestures(cls._scriptDecoratorGestures)
		except AttributeError:
			pass
	super(ScriptableObject, self).__init__()'''

SCRIPTABLE_OBJECT_BIND_GESTURE = r'''
def bindGesture(self, gestureIdentifier, scriptName):
	"""Bind an input gesture to a script.
	@param gestureIdentifier: The identifier of the input gesture.
	@type gestureIdentifier: str
	@param scriptName: The name of the script, which is the name of the method excluding the C{script_} prefix.
	@type scriptName: str
	@raise LookupError: If there is no script with the provided name.
	"""
	scriptAttrName = "script_%s" % scriptName
	# Don't store the instance method, as this causes a circular reference
	# and instance methods are meant to be generated on retrieval anyway.
	func = getattr(self.__class__, scriptAttrName, None)
	if not func:
		raise LookupError(
			"No such script on class {className}. Couldn't find attribute: {scriptAttrName}".format(
				className=self.__class__.__name__,
				scriptAttrName=scriptAttrName,
			),
		)
	# Import late to avoid circular import.
	import inputCore

	self._gestureMap[inputCore.normalizeGestureIdentifier(gestureIdentifier)] = func'''

SCRIPTABLE_OBJECT_REMOVE_GESTURE_BINDING = r'''
def removeGestureBinding(self, gestureIdentifier):
	"""
	Removes the binding for the given gesture identifier if a binding exists.
	@param gestureIdentifier: The identifier of the input gesture.
	@type gestureIdentifier: str
	@raise LookupError: If there is no binding for this gesture
	"""
	# Import late to avoid circular import.
	import inputCore

	del self._gestureMap[inputCore.normalizeGestureIdentifier(gestureIdentifier)]'''

SCRIPTABLE_OBJECT_BIND_GESTURES = r'''
def bindGestures(self, gestureMap):
	"""Bind or unbind multiple input gestures.
	This is a convenience method which simply calls L{bindGesture} for each gesture and script pair, logging any errors.
	For the case where script is None, L{removeGestureBinding} is called instead.
	@param gestureMap: A mapping of gesture identifiers to script names.
	@type gestureMap: dict of str to str
	"""
	for gestureIdentifier, scriptName in gestureMap.items():
		if scriptName:
			try:
				self.bindGesture(gestureIdentifier, scriptName)
			except LookupError:
				log.error("Error binding script %s in %r" % (scriptName, self))
		else:
			try:
				self.removeGestureBinding(gestureIdentifier)
			except LookupError:
				pass'''

SCRIPTABLE_OBJECT_GET_SCRIPT = r'''
def getScript(self, gesture):
	"""Retrieve the script bound to a given gesture.
	@param gesture: The input gesture in question.
	@type gesture: L{inputCore.InputGesture}
	@return: The script function or C{None} if none was found.
	@rtype: script function
	"""
	for identifier in gesture.normalizedIdentifiers:
		try:
			# Convert to instance method.
			return self._gestureMap[identifier].__get__(self, self.__class__)
		except KeyError:
			continue
		except AttributeError:
			log.exception(
				f"Base class may not have been initialized.\nMRO={self.__class__.__mro__}"
				if not hasattr(self, "_gestureMap")
				else None,
			)
			return None
	else:
		return None'''

GLOBAL_PLUGIN = r'''
class GlobalPlugin(baseObject.ScriptableObject):
	"""Base global plugin.
	Global plugins facilitate the implementation of new global commands,
	support for objects which may be found across many applications, etc.
	Each global plugin should be a separate Python module in the globalPlugins package containing a C{GlobalPlugin} class which inherits from this base class.
	Global plugins can implement and bind gestures to scripts which will take effect at all times.
	See L{ScriptableObject} for details.
	Global plugins can also receive NVDAObject events for all NVDAObjects.
	This is done by implementing methods called C{event_eventName},
	where C{eventName} is the name of the event; e.g. C{event_gainFocus}.
	These event methods take two arguments: the NVDAObject on which the event was fired
	and a callable taking no arguments which calls the next event handler.
	"""

	def terminate(self):
		"""Terminate this global plugin.
		This will be called when NVDA is finished with this global plugin.
		"""

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		"""Choose NVDAObject overlay classes for a given NVDAObject.
		This is called when an NVDAObject is being instantiated after L{NVDAObjects.NVDAObject.findOverlayClasses} has been called on the API-level class.
		This allows a global plugin to add or remove overlay classes.
		See L{NVDAObjects.NVDAObject.findOverlayClasses} for details about overlay classes.
		@param obj: The object being created.
		@type obj: L{NVDAObjects.NVDAObject}
		@param clsList: The list of classes, which will be modified by this method if appropriate.
		@type clsList: list of L{NVDAObjects.NVDAObject}
		"""

	def __repr__(self):
		return f"{self.__class__.__name__} ({self.__class__.__module__!r})"'''

SCRIPT_DECORATOR = r'''
def script(
	description: str = "",
	category: Optional[str] = None,
	gesture: Optional[str] = None,
	gestures: Optional[Iterator[str]] = None,
	canPropagate: bool = False,
	bypassInputHelp: bool = False,
	allowInSleepMode: bool = False,
	resumeSayAllMode: Optional[int] = None,
	speakOnDemand: bool = False,
):
	"""Define metadata for a script.
	This function is to be used as a decorator to set metadata used by the scripting system and gesture editor.
	It can only decorate methods which have a name starting with "script_"
	:param description: A short translatable description of the script to be used in the gesture editor, etc.
	:param category: The category of the script displayed in the gesture editor.
	:param gesture: A gesture associated with this script.
	:param gestures: A collection of gestures associated with this script
	:param canPropagate: Whether this script should also apply when it belongs to a  focus ancestor object.
	:param bypassInputHelp: Whether this script should run when input help is active.
	:param allowInSleepMode: Whether this script should run when NVDA is in sleep mode.
	:param resumeSayAllMode: The say all mode that should be resumed when active before executing this script.
	One of the C{sayAll.CURSOR_*} constants.
	:param speakOnDemand: Whether this script should speak when NVDA speech mode is "on-demand"
	"""
	if gestures is None:
		gestures: List[str] = []
	else:
		# A tuple may have been used, however, the collection of gestures may need to be
		# extended (via append) with the value of the 'gesture' string (in-case both are provided in the
		# decorator).
		gestures: List[str] = list(gestures)

	def script_decorator(decoratedScript):
		# Decoratable scripts are functions, not bound instance methods.
		if not isinstance(decoratedScript, types.FunctionType):
			log.warning(
				"Using the script decorator is unsupported for %r" % decoratedScript,
				stack_info=True,
			)
			return decoratedScript
		if not decoratedScript.__name__.startswith("script_"):
			log.warning(
				"Can't apply  script decorator to %r which name does not start with 'script_'"
				% decoratedScript.__name__,
				stack_info=True,
			)
			return decoratedScript
		decoratedScript.__doc__ = description
		if category is not None:
			decoratedScript.category = category
		if gesture is not None:
			gestures.append(gesture)
		if gestures:
			decoratedScript.gestures = gestures
		decoratedScript.canPropagate = canPropagate
		decoratedScript.bypassInputHelp = bypassInputHelp
		if resumeSayAllMode is not None:
			decoratedScript.resumeSayAllMode = resumeSayAllMode
		decoratedScript.allowInSleepMode = allowInSleepMode
		decoratedScript.speakOnDemand = speakOnDemand
		return decoratedScript

	return script_decorator'''

GET_OBJ_SCRIPT = r'''
def _getObjScript(
	obj: "NVDAObjects.NVDAObject",
	gesture: "inputCore.InputGesture",
	globalMapScripts: List["inputCore.InputGestureScriptT"],
) -> Optional[_ScriptFunctionT]:
	"""
	@param globalMapScripts: An ordered list of scripts.
	The list is ordered by resolution priority,
	the first map in the list should be used to resolve the script first.
	"""
	# Search the scripts from the global gesture maps.
	for cls, scriptName in globalMapScripts:
		if isinstance(obj, cls):
			if scriptName is None:
				# The global map specified that no script should execute for this gesture and object.
				return None
			if scriptName.startswith("kb:"):
				# Emulate a key press.
				return _makeKbEmulateScript(scriptName)
			try:
				return getattr(obj, "script_%s" % scriptName)
			except AttributeError:
				pass

	try:
		# Search the object itself for in-built bindings.
		return obj.getScript(gesture)
	except Exception:  # Prevent a faulty add-on from breaking script handling altogether (#5446)
		log.exception()'''

NORMALIZE_GESTURE_IDENTIFIER = r'''
def normalizeGestureIdentifier(identifier):
	"""Normalize a gesture identifier so that it matches other identifiers for the same gesture.
	First, the entire identifier is converted to lower case.
	Then, any items separated by a + sign after the source prefix are considered to be of indeterminate order
	and are sorted by character.
	This is done because, for example, "kb:shift+alt+downArrow"
	must be treated the same as "kb:alt+shift+downarrow".
	"""
	identifier = identifier.lower()
	prefix, main = identifier.split(":", 1)
	main = main.split("+")
	# The order of the parts doesn't matter as far as the user is concerned,
	# but we need them to be in a determinate order so they will match other gesture identifiers.
	# We sort them by character.
	main.sort()
	main = "+".join(main)
	return "{0}:{1}".format(prefix, main)'''

SCRIPTABLE_OBJECT = "class ScriptableObject(AutoPropertyObject, metaclass=ScriptableType):\n" + "".join(
	textwrap.indent(method, "\t")
	for method in (
		SCRIPTABLE_OBJECT_INIT,
		SCRIPTABLE_OBJECT_BIND_GESTURE,
		SCRIPTABLE_OBJECT_REMOVE_GESTURE_BINDING,
		SCRIPTABLE_OBJECT_BIND_GESTURES,
		SCRIPTABLE_OBJECT_GET_SCRIPT,
	)
)



class _StubModule(types.ModuleType):
	"""An NVDA module whose every attribute is a mock, unless the test sets it."""

	def __getattr__(self, name):
		if name.startswith("__"):
			raise AttributeError(name)
		value = mock.MagicMock(name="%s.%s" % (self.__name__, name))
		setattr(self, name, value)
		return value


def _nvda():
	"""NVDA 2026.2's scripting code, run from the text above."""
	nvda = {
		"Any": typing.Any,
		"Iterator": typing.Iterator,
		"List": typing.List,
		"Optional": typing.Optional,
		"_ScriptFunctionT": typing.Callable,
		"types": types,
		"log": mock.MagicMock(name="log"),
		# AutoPropertyType and AutoPropertyObject only add NVDA's _get_ properties.
		"AutoPropertyType": type,
		"AutoPropertyObject": object,
	}
	for code in (NORMALIZE_GESTURE_IDENTIFIER, SCRIPTABLE_TYPE, SCRIPTABLE_OBJECT, SCRIPT_DECORATOR, GET_OBJ_SCRIPT):
		exec(code, nvda)
	nvda["baseObject"] = types.SimpleNamespace(ScriptableObject=nvda["ScriptableObject"])
	exec(GLOBAL_PLUGIN, nvda)
	return nvda


NVDA = _nvda()
_savedModules = {}
_hadTranslation = hasattr(builtins, "_")
plugin = None


def _stub(name, **attributes):
	module = _StubModule(name)
	for key, value in attributes.items():
		setattr(module, key, value)
	_savedModules.setdefault(name, sys.modules.get(name))
	sys.modules[name] = module
	if "." in name:
		parent, child = name.rsplit(".", 1)
		setattr(sys.modules[parent], child, module)
	return module


def setUpModule():
	global plugin
	for name in (
		"addonHandler", "api", "core", "globalVars", "keyboardHandler", "logHandler", "NVDAObjects",
		"NVDAObjects.IAccessible", "oleacc", "ui", "windowUtils", "winUser", "gui", "gui.guiHelper",
	):
		_stub(name)
	_stub("baseObject", ScriptableObject=NVDA["ScriptableObject"], ScriptableType=NVDA["ScriptableType"])
	_stub("inputCore", normalizeGestureIdentifier=NVDA["normalizeGestureIdentifier"])
	_stub("scriptHandler", script=NVDA["script"])
	_stub("globalPluginHandler", GlobalPlugin=NVDA["GlobalPlugin"], runningPlugins=set())
	_stub("gui.settingsDialogs", SettingsPanel=type("SettingsPanel", (), {}))
	_stub("wx", Dialog=type("Dialog", (), {}))
	if not _hadTranslation:
		builtins._ = lambda text: text
	spec = importlib.util.spec_from_file_location(
		"replyToSenderOutlook",
		os.path.join(PLUGIN_DIR, "__init__.py"),
		submodule_search_locations=[PLUGIN_DIR],
	)
	plugin = importlib.util.module_from_spec(spec)
	_savedModules.setdefault("replyToSenderOutlook", sys.modules.get("replyToSenderOutlook"))
	sys.modules["replyToSenderOutlook"] = plugin
	spec.loader.exec_module(plugin)


def tearDownModule():
	for name in list(sys.modules):
		if name.startswith("replyToSenderOutlook.") and name not in _savedModules:
			del sys.modules[name]
	for name, module in _savedModules.items():
		if module is None:
			sys.modules.pop(name, None)
		else:
			sys.modules[name] = module
	if not _hadTranslation:
		del builtins._


def _running(pluginClass):
	"""A running global plugin, without GlobalPlugin.__init__'s update check."""
	instance = object.__new__(pluginClass)
	NVDA["ScriptableObject"].__init__(instance)
	return instance


def _keyPress(identifier):
	"""A keyboard gesture as NVDA's KeyboardInputGesture names it: with and without the layout."""
	main = identifier.split(":", 1)[1]
	return types.SimpleNamespace(
		normalizedIdentifiers=[
			NVDA["normalizeGestureIdentifier"]("kb(laptop):" + main),
			NVDA["normalizeGestureIdentifier"]("kb:" + main),
		],
	)


class OutlookExtendedAppModule(object):
	"""Stands in for Outlook Extended's AppModule, which builds on NVDA's and so has nativeOm."""

	appName = "outlook"

	def __init__(self, nativeOm):
		self.nativeOm = nativeOm


def _inApp(appName, nativeOm=None):
	appModule = OutlookExtendedAppModule(nativeOm)
	appModule.appName = appName
	return types.SimpleNamespace(appModule=appModule, windowHandle=0x1234)


def _focusIn(appName, nativeOm=None):
	return mock.patch.object(plugin.api, "getFocusObject", return_value=_inApp(appName, nativeOm))


def _findScript(plugins, gesture, globalMapScripts=()):
	"""The global plugin level of NVDA's scriptHandler.findScript."""
	for running in plugins:
		script = NVDA["_getObjScript"](running, gesture, list(globalMapScripts))
		if script:
			return script
	return None


OUTLOOK_KEYS = (
	("kb:NVDA+shift+r", "replyToSender"),
	("kb:NVDA+shift+e", "announceSender"),
)


class KeyTests(unittest.TestCase):
	def setUp(self):
		self.plugin = _running(plugin.GlobalPlugin)

	def test_keysAreBound(self):
		self.assertIs(self.plugin._gestureMap.get("kb:nvda+r+shift"), plugin.GlobalPlugin.script_replyToSender)
		self.assertIs(self.plugin._gestureMap.get("kb:e+nvda+shift"), plugin.GlobalPlugin.script_announceSender)
		for name in ("replyToSender", "announceSender", "checkForUpdates"):
			self.assertEqual(getattr(plugin.GlobalPlugin, "script_" + name).category, "Reply to Sender for Outlook")

	def test_keysWorkInClassicOutlook(self):
		om = object()
		for key, function in OUTLOOK_KEYS:
			with self.subTest(key=key), _focusIn("outlook", om), mock.patch.object(plugin.outlookReply, function) as work:
				script = _findScript([self.plugin], _keyPress(key))
				self.assertIsNotNone(script)
				script(None)
			# Outlook's object model comes from whichever Outlook app module NVDA runs.
			work.assert_called_once_with(om)

	def test_keysAreLeftToOthersOutsideClassicOutlook(self):
		class OtherAddon(NVDA["GlobalPlugin"]):
			@NVDA["script"](gestures=["kb:NVDA+shift+r", "kb:NVDA+shift+e"])
			def script_other(self, gesture):
				pass

		other = _running(OtherAddon)
		# olk is the new Outlook, which has no object model to reply through.
		for appName in ("notepad", "chrome", "explorer", "olk"):
			for key, function in OUTLOOK_KEYS:
				with self.subTest(appName=appName, key=key), _focusIn(appName):
					self.assertIsNone(_findScript([self.plugin], _keyPress(key)))
					script = _findScript([self.plugin, other], _keyPress(key))
					self.assertIs(script.__func__, OtherAddon.script_other)

	def test_updateCheckIsNotLimitedToOutlook(self):
		self.plugin.bindGesture("kb:NVDA+shift+f9", "checkForUpdates")
		with _focusIn("notepad"):
			script = _findScript([self.plugin], _keyPress("kb:NVDA+shift+f9"))
		self.assertIs(script.__func__, plugin.GlobalPlugin.script_checkForUpdates)

	def test_keyAssignedInInputGesturesOutsideOutlook(self):
		# A key the user assigns is in NVDA's gesture map and reaches the script anywhere.
		for key, function in OUTLOOK_KEYS:
			globalMap = [(plugin.GlobalPlugin, function)]
			with self.subTest(key=key), _focusIn("notepad"), mock.patch.object(
				plugin.outlookReply, function,
			) as work, mock.patch.object(plugin.ui, "message") as message:
				script = _findScript([self.plugin], _keyPress("kb:NVDA+shift+f5"), globalMap)
				self.assertIs(script.__func__, getattr(plugin.GlobalPlugin, "script_" + function))
				script(None)
			work.assert_not_called()
			message.assert_called_once_with("No message to reply to. Open or select a message in classic Outlook.")


class ForegroundTests(unittest.TestCase):
	def setUp(self):
		self.plugin = _running(plugin.GlobalPlugin)
		self.nextHandler = mock.MagicMock(name="nextHandler")

	def test_messageWindowsInOutlookAreMaximized(self):
		om = object()
		obj = _inApp("outlook", om)
		with mock.patch.object(plugin.outlookReply, "maximizeIfInspector") as maximize:
			self.plugin.event_foreground(obj, self.nextHandler)
		maximize.assert_called_once_with(om, obj)
		self.nextHandler.assert_called_once_with()

	def test_otherProgramsAreLeftAlone(self):
		obj = _inApp("notepad")
		with mock.patch.object(plugin.outlookReply, "maximizeIfInspector") as maximize:
			self.plugin.event_foreground(obj, self.nextHandler)
		maximize.assert_not_called()
		self.nextHandler.assert_called_once_with()

	def test_otherOutlookAddonsStillGetTheEvent(self):
		with mock.patch.object(plugin.outlookReply, "maximizeIfInspector", side_effect=RuntimeError):
			self.plugin.event_foreground(_inApp("outlook", object()), self.nextHandler)
		self.nextHandler.assert_called_once_with()


class MaximizeTests(unittest.TestCase):
	"""outlookReply.maximizeIfInspector, which 0.3 ran from its app module's event_foreground."""

	def setUp(self):
		self.user32 = mock.MagicMock(name="user32", **{"IsZoomed.return_value": 0})
		patcher = mock.patch.object(plugin.outlookReply.ctypes, "windll", types.SimpleNamespace(user32=self.user32))
		patcher.start()
		self.addCleanup(patcher.stop)
		winUser = plugin.outlookReply.winUser
		winUser.getClassName = mock.MagicMock(return_value="rctrl_renwnd32")
		winUser.getWindowText = mock.MagicMock(return_value="Hello - Message (HTML)")
		self.om = mock.MagicMock(name="nativeOm")
		self.om.ActiveInspector.return_value.Caption = "Hello - Message (HTML)"

	def test_messageWindowIsMaximized(self):
		plugin.outlookReply.maximizeIfInspector(self.om, _inApp("outlook"))
		self.user32.ShowWindow.assert_called_once_with(0x1234, 3)

	def test_mainWindowIsNot(self):
		plugin.outlookReply.winUser.getWindowText.return_value = "Inbox - josh@example.com - Outlook"
		plugin.outlookReply.maximizeIfInspector(self.om, _inApp("outlook"))
		self.user32.ShowWindow.assert_not_called()

	def test_noObjectModel(self):
		plugin.outlookReply.maximizeIfInspector(None, _inApp("outlook"))
		self.user32.ShowWindow.assert_not_called()


class ReplyTests(unittest.TestCase):
	"""The reply itself, moved unchanged from 0.3's app module into outlookReply."""

	def setUp(self):
		self.item = mock.MagicMock(name="item", Class=43)
		self.item.PropertyAccessor.GetProperty.return_value = "sender@example.com"
		self.om = mock.MagicMock(name="nativeOm")
		self.om.ActiveWindow.return_value.CurrentItem = self.item
		for name in ("api", "core", "ui"):
			patcher = mock.patch.object(plugin.outlookReply, name)
			patcher.start()
			self.addCleanup(patcher.stop)
		plugin.outlookReply.winUser.getClassName = mock.MagicMock(return_value="Other")

	def test_replyGoesToTheSenderOnly(self):
		plugin.outlookReply.replyToSender(self.om)
		reply = self.item.Reply.return_value
		self.assertEqual(reply.To, "sender@example.com")
		self.assertEqual((reply.CC, reply.BCC), ("", ""))
		reply.Recipients.ResolveAll.assert_called_once_with()
		reply.Display.assert_called_once_with(False)
		plugin.outlookReply.api.copyToClip.assert_called_once_with("sender@example.com")
		plugin.outlookReply.ui.message.assert_not_called()

	def test_announceSender(self):
		plugin.outlookReply.announceSender(self.om)
		plugin.outlookReply.ui.message.assert_called_once_with("sender@example.com")

	def test_noObjectModel(self):
		plugin.outlookReply.replyToSender(None)
		plugin.outlookReply.ui.message.assert_called_once_with("Unable to connect to Outlook.")


if __name__ == "__main__":
	unittest.main()
