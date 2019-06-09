# AutoFillWordForms
A way that I can automatically fill in word forms on the Mac.

I used kotlin as it works on the JVM and Apache POI has a bunch of helper methods of parsing out word forms.

I created an automator task to be able to call the java jar program with a alright file dialog UI (without any work on this part)

# Things I learned

A form field may span multiple runs in multiple paragraphs.  As such multiple text entries can be applied to the same form field.

Basic is:
FormField Begin
runs (for spans of text with markup applied to it)
paragraphs (like what happens by default when user makes a new line)
FormField End

# TODO
- handle date fields better
