#ifndef __INITCOMMONCONTROLS_H__83716747_6B45_4c0a_8B18_81C4751EA455_
#define __INITCOMMONCONTROLS_H__83716747_6B45_4c0a_8B18_81C4751EA455_

#include <commctrl.h>
#pragma comment(lib, "comctl32.lib")

/// <Summary>
/// Helper Class to ensure the initialization of Common Controls Library. 
/// Add this as a Base Class for Your Control Class to Initialize the Common Control Library Automatically.
/// 
/// Since this Class uses a local static variable in the constructor for the Library Initialization,
/// the InitCommonControlsEx() is called only once per App for all the derived class objects.
///
/// Being a base class - this guarantees that the library is initialized before any derived class
/// functions are invoked.
///
/// In case your control has multiple base classes, try to make the CInitCommonControls as the
/// first base class followed by the others. This makes sure that InitCommonControlsEx() is called 
/// before any other dependent functions in other base classes are invoked.
/// </Summary>
/// <remarks>
/// Usage:
///	class MyClass: CInitCommonControls<ICC_LISTVIEW_CLASSES | ICC_TREEVIEW_CLASSES>
///	{
///		//...
///	}
/// </remarks>

template<DWORD dwICC = ICC_WIN95_CLASSES>
class CInitCommonControls
{

protected:

	/// Protected Constructor
	inline CInitCommonControls(void)
	{
		///
		/// Helper Structure to Call the InitCommonControlsEx() function
		///
		static struct _tagInitCommonControls
		{
			_tagInitCommonControls()
			{
				INITCOMMONCONTROLSEX icce;
				icce.dwSize = sizeof(INITCOMMONCONTROLSEX);
				icce.dwICC	= dwICC; // Use the Template Argument for the Library Initialization
				InitCommonControlsEx(&icce);
			}
		} DummyInitCommonControls;	// Static Local Variable Gets initialized only once (for the first Object)
	}

	/// Protected Destructor
	inline ~CInitCommonControls(void)
	{
	}

};

#endif