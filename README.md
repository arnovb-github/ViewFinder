# ViewFinder
Commence RM script for finding views in the database.

# What does it do?
View Finder will display a list of all Views in a Commence database. You can search for views by typing a few characters of the view name. You can also filter views based on the category they belong to. Or both. Double-click the resulting views to open them.

# How is that useful?
When you have a lot of views in a Commence database, finding the right one can be a bit awkward. Unless you have them at your fingertips in the View Bar of View Menu, you have to use the menu option View | Open/Manage views. This will give you a list of all views present in the database. Once you drill down to a view, you can use the Open button (or press Enter) to open the view. If it happens not to be the one you wanted, you have to do this all over again. ViewFinder can be quite a time-saver.

# System requirements
Commence RM 3.1 or higher.

# Usage
Create a new Detail Form in any category. Tip: I use a dedicated one called 'Scripts', specifically for custom detail form scripts like this one. Create a layout as you see fit (see screenshots for example). Be sure to include the required controls. The form does not use any bound fields (=database fields). Check in the form script. Create an Agent that performs an Add Item, and displays this Detail Form.

Controls used in the script: TextBox, ComboBox, CommandButton, ListView (ActiveX control), StatusBar (ActiveX control). Be sure to include them on the form. Putting ActiveX controls on a form will result in many error messages. Simply ignore them.

I've included the source XML of the form script, If you are really handy with Commence it may come in handy.

# Limitations
MultiViews are not exposed by the Commence API and cannot be searched for.

# Note
This GIT was created just for archiving purposes. Do not expect updates.
