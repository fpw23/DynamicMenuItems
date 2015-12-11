# DynamicMenuItems
Demo showing how to attach dynamic menu items list to a context menu in visual studio

There is a [demo](https://msdn.microsoft.com/en-us/library/bb166492.aspx) of how to do this provided by Microsoft but it is not that clear so I am sharing my enhanced version I built while researching this topic.  My needs were a little different, I wanted to be able to dynamically hide show a context menu that list dynamic menu items based on how the menu was invoke.  

To use, download the code and build the project and start debugging.  Once you have VS dev env running, open any project and create a folder called Classes.  Right click on this folder and you should see a new menu item called "Run T4 Item...".  The menu items in the list were created dynamically.  If you click on one, it will pop a message box and add a new project scoped dynamic menu.  Right click on the poject and you will see the same "Run T4 Item..." menu.

While the Microsoft demo did work, it did not make sense.  There is no explaination of how the items were added to the menu list.  It took me a while to understand what was happening but now I get it and I think this example might make it easier for others.

Here is the way I understand dynamic menus to work.

You create the menus and buttons in the vsct file like normal except for the dynamic one, for that one you have to add the command flag:

       <CommandFlag>DynamicItemStart</CommandFlag>

this signals to visual studio that your menu will be dynamic. 

You build a class that will manage what is shown on the menu.  From the ms example they built a class called DynamicMenu and its job is to register the dynamic menu button and provide the three methods vs will call into when it encounters the button:  OnBeforeQueryStatusDynamicItem, IsValidDynamicItem, OnInvokedDynamicItem.  When you right click and invoke the context menu, vs hits your dynamic button and then calls these methods over and over allow you to set as many menu items as you need.  The IsValidDynamicItem method controls the number of times by returning false when you are done.   The OnBeforeQueryStatusDynamicItem lets you set the menu text for the button. And the OnInvokedDynamicItem lets you respond the to user clicking the menu item.


