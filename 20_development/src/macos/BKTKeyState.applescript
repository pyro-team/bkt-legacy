use framework "AppKit"
use scripting additions

on modifierFlags(paramString)
	set flagsValue to (current application's NSEvent's modifierFlags()) as integer
	return flagsValue as text
end modifierFlags
