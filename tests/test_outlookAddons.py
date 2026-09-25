"""Reply to Sender must leave other Outlook add-ons working (0.4).

python -m unittest discover -s tests

Reported in github.com/joshknnd1982/outlookFirstLineSilence issue #4: with
Reply to Sender 0.3 installed, Outlook Extended's commands (Alt+1 and so on for
a message's header fields) did nothing, and the attached log showed Reply to
Sender's appModules/outlook.py handling Outlook's events.

NVDA runs one app module for Outlook, appModules.outlook: the first it finds on
appModules.__path__. Each add-on with an appModules folder goes to the front of
that path (NVDA 2026.2's Addon.addToPackagePath, copied word for word below),
in the order the log shows, so an add-on named after outlookExtended hides it.
Global plugins are different: NVDA 2026.2's listPlugins, also below, loads one
from every add-on. The rest of NVDA is stubbed.
"""

import importlib
import importlib.util
import os
import pkgutil
import shutil
import sys
import tempfile
import types
import unittest
import zipfile
from unittest import mock

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ADDON_DIR = os.path.join(ROOT, "addon")

# NVDA 2026.2 code, word for word: addonHandler/__init__.py, class Addon.
ADD_TO_PACKAGE_PATH = r'''
def addToPackagePath(self, package):
	"""Adds this L{Addon} extensions to the specific package path if those exist.
	This allows the addon to "run" / be available because the package is able to search its path,
	looking for particular modules. This is used by the following:
	- `globalPlugins`
	- `appModules`
	- `synthDrivers`
	- `brailleDisplayDrivers`
	@param package: the python module representing the package.
	@type package: python module.
	"""
	# #3090: Ensure that we don't add disabled / blocked add-ons to package path.
	# By returning here the addon does not "run"/ become active / registered.
	if self.isDisabled or self.isBlocked or self.isPendingInstall or self.name in _failedPendingRemovals:
		return

	extension_path = os.path.join(self.path, package.__name__)
	if not os.path.isdir(extension_path):
		# This addon does not have extension points for this package
		return
	converted_path = self._getPathForInclusionInPackage(package)
	package.__path__.insert(0, converted_path)
	self._extendedPackages.add(package)
	log.debug("Addon %s added to %s package path", self.manifest["name"], package.__name__)'''

GET_PATH_FOR_INCLUSION_IN_PACKAGE = r'''
def _getPathForInclusionInPackage(self, package):
	extension_path = os.path.join(self.path, package.__name__)
	return extension_path'''

# NVDA 2026.2 code, word for word: globalPluginHandler.py.
LIST_PLUGINS = r'''
def listPlugins():
	for loader, name, isPkg in pkgutil.iter_modules(globalPlugins.__path__):
		if name.startswith("_"):
			continue
		try:
			plugin = importlib.import_module("globalPlugins.%s" % name, package="globalPlugins").GlobalPlugin
		except:  # noqa: E722
			log.error("Error importing global plugin %s" % name, exc_info=True)
			continue
		yield plugin'''

#: The add-ons with Outlook code in Dennis's log for issue #4, in the order NVDA added them.
DENNIS_ORDER = ("outlookExtended", "outlookFirstLineSilence", "replyToSenderOutlook")


def _nvda():
	nvda = {"os": os, "pkgutil": pkgutil, "importlib": importlib, "_failedPendingRemovals": set()}
	nvda["log"] = mock.MagicMock(name="log")
	for code in (ADD_TO_PACKAGE_PATH, GET_PATH_FOR_INCLUSION_IN_PACKAGE, LIST_PLUGINS):
		exec(code, nvda)
	return nvda


NVDA = _nvda()


class Addon(object):
	"""An installed, enabled add-on, with NVDA's own addToPackagePath."""

	isDisabled = isBlocked = isPendingInstall = False
	addToPackagePath = NVDA["addToPackagePath"]
	_getPathForInclusionInPackage = NVDA["_getPathForInclusionInPackage"]

	def __init__(self, path):
		self.path = path
		self.name = os.path.basename(path)
		self.manifest = {"name": self.name}
		self._extendedPackages = set()


def _write(path, text):
	os.makedirs(os.path.dirname(path), exist_ok=True)
	with open(path, "w", encoding="utf-8") as f:
		f.write(text)


class OutlookAppModuleTests(unittest.TestCase):
	def setUp(self):
		self.tmp = tempfile.mkdtemp()
		self.addCleanup(shutil.rmtree, self.tmp, True)
		self.addCleanup(self._forgetPackages)
		self._forgetPackages()
		# NVDA's own Outlook support is last on the path, after every add-on.
		_write(os.path.join(self.tmp, "nvda", "appModules", "outlook.py"), 'WHO = "NVDA"\n')
		# Outlook Extended 3.4 has an appModules/outlook package, built on NVDA's.
		_write(
			os.path.join(self.tmp, "addons", "outlookExtended", "appModules", "outlook", "__init__.py"),
			'WHO = "Outlook Extended"\n',
		)
		_write(
			os.path.join(self.tmp, "addons", "outlookFirstLineSilence", "globalPlugins", "outlookFirstLineSilence", "__init__.py"),
			"class GlobalPlugin(object):\n\tpass\n",
		)
		self.replyToSender = os.path.join(self.tmp, "addons", "replyToSenderOutlook")
		shutil.copytree(ADDON_DIR, self.replyToSender, ignore=shutil.ignore_patterns("__pycache__"))

	@staticmethod
	def _forgetPackages():
		for name in list(sys.modules):
			if name.split(".")[0] in ("appModules", "globalPlugins"):
				del sys.modules[name]
		importlib.invalidate_caches()

	def _package(self, name, builtinPath):
		package = types.ModuleType(name)
		package.__path__ = [builtinPath]
		sys.modules[name] = package
		for addonName in DENNIS_ORDER:
			Addon(os.path.join(self.tmp, "addons", addonName)).addToPackagePath(package)
		return package

	def _outlookAppModule(self):
		self._package("appModules", os.path.join(self.tmp, "nvda", "appModules"))
		return importlib.import_module("appModules.outlook", package="appModules")

	def test_outlookExtendedIsTheOutlookAppModule(self):
		self.assertEqual(self._outlookAppModule().WHO, "Outlook Extended")

	def test_nvdasOwnWithoutOutlookExtended(self):
		shutil.rmtree(os.path.join(self.tmp, "addons", "outlookExtended"))
		self.assertEqual(self._outlookAppModule().WHO, "NVDA")

	def test_addonHasNoAppModules(self):
		self.assertFalse(os.path.exists(os.path.join(ADDON_DIR, "appModules")))

	def test_anAppModuleLikeVersion03HidesOutlookExtended(self):
		# What went wrong: 0.3 had appModules/outlook.py, and NVDA imported it instead.
		_write(os.path.join(self.replyToSender, "appModules", "outlook.py"), 'WHO = "Reply to Sender 0.3"\n')
		self.assertEqual(self._outlookAppModule().WHO, "Reply to Sender 0.3")

	def test_everyOutlookAddonsGlobalPluginLoads(self):
		# listPlugins imports globalPlugins.<name> for every add-on; the imports are recorded, not run.
		NVDA["globalPlugins"] = self._package("globalPlugins", os.path.join(self.tmp, "nvda", "globalPlugins"))
		def importModule(name, package=None):
			return types.SimpleNamespace(GlobalPlugin=name)

		with mock.patch.dict(NVDA, importlib=types.SimpleNamespace(import_module=importModule)):
			plugins = list(NVDA["listPlugins"]())
		self.assertIn("globalPlugins.outlookFirstLineSilence", plugins)
		self.assertIn("globalPlugins.replyToSenderOutlook", plugins)

	def test_builtPackageHasNoAppModules(self):
		with open(os.path.join(ADDON_DIR, "manifest.ini"), encoding="utf-8") as f:
			version = [line.split("=", 1)[1].strip() for line in f if line.startswith("version")][0]
		built = os.path.join(ROOT, "replyToSenderOutlook-%s.nvda-addon" % version)
		if not os.path.exists(built):
			self.skipTest("run python build.py first")
		with zipfile.ZipFile(built) as bundle:
			names = bundle.namelist()
		self.assertIn("globalPlugins/replyToSenderOutlook/outlookReply.py", names)
		self.assertEqual([name for name in names if name.startswith("appModules/")], [])


if __name__ == "__main__":
	unittest.main()
