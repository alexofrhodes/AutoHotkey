/*
	Function: Anchor (translated to v2 by Relayer 2023-02-22)
		Defines how controls should be automatically positioned relative to the new dimensions of a window when resized.

	Parameters:
		i - a control HWND, associated variable name or ClassNN to operate on
		a - (optional) one or more of the anchors: 'x', 'y', 'w' (width) and 'h' (height),
			optionally followed by a relative factor, e.g. "x h0.5"
		r - (optional) true to redraw controls, recommended for GroupBox and Button types

	Examples:
> "xy" ; bounds a control to the bottom-left edge of the window
> "w0.5" ; any change in the width of the window will resize the width of the control on a 2:1 ratio
> "h" ; similar to above but directly proportional to height

	Remarks:
		To assume the current window size for the new bounds of a control (i.e. resetting) simply omit the second and third parameters.
		However if the control had been created with DllCall() and has its own parent window,
			the container AutoHotkey created GUI must be made default with the +LastFound option prior to the call.
		For a complete example see anchor-example.ahk.

	License:
		- Version 4.60a <http://www.autohotkey.net/~Titan/#anchor>
		- New BSD License <http://www.autohotkey.net/~Titan/license.txt>
*/
;64 bit compatible
Anchor(i, a := "", r := false)
{
	static c, cs := 12, cx := 255, cl := 0, g, gs := 8, gl := 0, gpi := "", gw, gh, z := 0, ptr
	cf := gf := 0
	if z = 0
	{
		g := Buffer(gs * 99, 0), c := Buffer(cs * cx, 15)
		ptr := A_PtrSize ? "Ptr" : "UInt", z := true
	}
	gi := Buffer(68, 0)
	DllCall("GetWindowInfo", "UInt", gp := DllCall("GetParent", "UInt", i), ptr, gi) ; &gi
		, giw := NumGet(gi, 28, "Int") - NumGet(gi, 20, "Int"), gih := NumGet(gi, 32, "Int") - NumGet(gi, 24, "Int")

	if (gp != gpi)
	{
		gpi := gp
		Loop gl
			if NumGet(g, cb := gs * (A_Index - 1), "UInt") == gp
			{
				gw := NumGet(g, cb + 4, "Short"), gh := NumGet(g, cb + 6, "Short"), gf := 1
				break
			}
		if !gf
			NumPut("UInt", gp, g, gl), NumPut("Short", gw := giw, g, gl + 4), NumPut("Short", gh := gih, g, gl + 6), gl += gs
	}
	ControlGetPos &dx, &dy, &dw, &dh, i

	Loop cl
		if NumGet(c, cb := cs * (A_Index - 1), "UInt") == i
		{
			if (a = "")
			{
				cf := 1
				break
			}
			giw -= gw, gih -= gh, ass := 1, dx := NumGet(c, cb + 4, "Short"), dy := NumGet(c, cb + 6, "Short")
				, cw := dw, dw := NumGet(c, cb + 8, "Short"), ch := dh, dh := NumGet(c, cb + 10, "Short")

			d := map( "x", dx
					, "y", dy
					, "w", dw
					, "h", dh )

			Loop Parse, a, "xywh"
				if A_Index > 1
				{
					av := SubStr(a, ass, 1), ass += 1 + StrLen(A_LoopField)
					d[av] += (InStr("yh", av) ? gih : giw) * (isNumber(A_LoopField) ? A_LoopField : 1)
				}
			dx := d["x"], dy := d["y"], dw := d["w"], dh := d["h"]
			DllCall("SetWindowPos", "UInt", i, "UInt", 0, "Int", dx, "Int", dy
				, "Int", InStr(a, "w") ? dw : cw, "Int", InStr(a, "h") ? dh : ch, "Int", 4)
			if r != 0
				DllCall("RedrawWindow", "UInt", i, "UInt", 0, "UInt", 0, "UInt", 0x0101) ; RDW_UPDATENOW | RDW_INVALIDATE
			return
		}

	if cf != 1
		cb := cl, cl += cs
	if cf = 1
		dw -= giw - gw, dh -= gih - gh
	NumPut("UInt", i, c, cb)
	NumPut("Short", dx, c, cb + 4)
	NumPut("Short", dy, c, cb + 6)
	NumPut("Short", dw, c, cb + 8)
	NumPut("Short", dh, c, cb + 10)
	return true
}