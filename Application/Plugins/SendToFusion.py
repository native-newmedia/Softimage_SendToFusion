# SendToFusion
#
# Tim Crowson for Magnetic Dreams
# February, 2013
#
# DESCRIPTION:
# Send Nulls and Cameras To Fusion, app-link style
#
# REQUIREMENTS:
# - PeyeonScript
# - A running instance of Fusion
#
# SUPPORTED PARAMETERS:
# 	- Translation X, Y, Z
#	- Rotation X, Y, Z
#	- Camera Field of View
#	- Camera Focal Length
#	- Camera Film Back Height and Width


import os
import win32com.client
from win32com.client import constants

null = None
false = 0
true = 1

def XSILoadPlugin( in_reg ):
	in_reg.Author = "Tim Crowson"
	in_reg.Name = "SendToFusion"
	in_reg.Major = 1
	in_reg.Minor = 0

	in_reg.RegisterCommand("SendToFusion","SendToFusion")
	in_reg.RegisterMenu(constants.siMenuSEObjectContextID,"SendToFusion_Menu",false,false)

	return true

def XSIUnloadPlugin( in_reg ):
	strPluginName = in_reg.Name
	return true

def SendToFusion_Init( in_ctxt ):
	oCmd = in_ctxt.Source
	oCmd.Description = "Send selected nulls and cameras directly to Fusion"
	oCmd.ReturnValue = true
	return true

def SendToFusion_Execute(  ):
	'''
	Send item data directly to Fusion, if Fusion is running
	'''

	import PeyeonScript
	fusion = PeyeonScript.scriptapp('Fusion')

	try:
		comp = fusion.GetCurrentComp()
	except:
		comp = None
		XSIUIToolkit.MsgBox('No comp found. Is Fusion running?', 48, 'Send to Fusion')

	if comp:
		sceneStart = int(Application.GetValue('PlayControl.In'))
		sceneEnd = int(Application.GetValue('PlayControl.Out'))

		def transferItem(item, regID):
			# First, check all items in Fusion to see if our selected item exists. Create it if it doesn't.
			allTools = comp.GetToolList(False, regID)
			if item.Name in [allTools[1.0].Name for x in allTools]:
				Application.LogMessage("Send to Fusion: Found '%s', updating values..." %item.Name)
				tool = comp.FindTool(str(item.Name))
			else:
				Application.LogMessage("Send to Fusion: Creating new item in Fusion for '%s'..." %item.Name)
				if regID == 'Camera3D':
					tool = comp.Camera3D()
				elif regID == 'Locator3D':
					tool = comp.Locator3D()
				elif regID == 'LightSpot':
					pass
				tool.SetAttrs({'TOOLS_Name': str(item.Name)})

			# Set xform inputs as animated
			tool.Transform3DOp.Translate.X.ConnectTo(comp.BezierSpline({}))
			tool.Transform3DOp.Translate.Y.ConnectTo(comp.BezierSpline({}))
			tool.Transform3DOp.Translate.Z.ConnectTo(comp.BezierSpline({}))
			tool.Transform3DOp.Rotate.X.ConnectTo(comp.BezierSpline({}))
			tool.Transform3DOp.Rotate.Y.ConnectTo(comp.BezierSpline({}))
			tool.Transform3DOp.Rotate.Z.ConnectTo(comp.BezierSpline({}))

			# Set some static parameters
			if regID == 'Camera3D':
				tool.AoV.ConnectTo(comp.BezierSpline({}))
				tool.FilmGate = 'User'
				tool.AovType = 1

			if regID == 'Locator3D':
				tool.MakeRenderable = 1

			# Config Progress Bar
			oBar = XSIUIToolkit.ProgressBar
			oBar.Maximum = ((sceneEnd - sceneStart) + 1)
			oBar.Step = 1
			oBar.Caption = 'Transferring %s...' %item.Name
			oBar.Visible = True

			# Transfer animation
			for frame in range((sceneEnd - sceneStart) + 1):
				fusionFrame = frame + 1
				xsiFrame = frame + sceneStart

				comp.CurrentTime = fusionFrame
				tool.Transform3DOp.Translate.X[fusionFrame] = Application.GetValue('%s.kine.global.posx' %item, xsiFrame)
				tool.Transform3DOp.Translate.Y[fusionFrame] = Application.GetValue('%s.kine.global.posy' %item, xsiFrame)
				tool.Transform3DOp.Translate.Z[fusionFrame] = Application.GetValue('%s.kine.global.posz' %item, xsiFrame)
				tool.Transform3DOp.Rotate.X[fusionFrame] = Application.GetValue('%s.kine.global.rotx' %item, xsiFrame)
				tool.Transform3DOp.Rotate.Y[fusionFrame] = Application.GetValue('%s.kine.global.roty' %item, xsiFrame)
				tool.Transform3DOp.Rotate.Z[fusionFrame] = Application.GetValue('%s.kine.global.rotz' %item, xsiFrame)

				# item-specific animation - Cameras
				if regID == 'Camera3D':
					tool.ApertureW = Application.GetValue("%s.camera.projplanewidth" %item, xsiFrame)
					tool.ApertureH = Application.GetValue("%s.camera.projplaneheight" %item, xsiFrame)
					tool.AoV[fusionFrame] = Application.GetValue("%s.camera.fov" %item, xsiFrame)
					tool.FLength = Application.GetValue("%s.camera.projplanedist" %item, xsiFrame)

				oBar.Increment()
			oBar.Visible = False

		# run the transfer function on selected items
		for item in Application.Selection:
			if item.Type == 'camera':
				transferItem(item, 'Camera3D')
			elif item.Type == 'null':
				transferItem(item, 'Locator3D')

	return true
	

def SendToFusion_Menu_Init( in_ctxt ):
	oMenu = in_ctxt.Source
	oMenu.AddCommandItem("Send To Fusion","SendToFusion")
	return true

