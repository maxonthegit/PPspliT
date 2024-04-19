```
    _____  _____           _ _ _______
   |  __ \|  __ \         | (_)__   __|
   | |__) | |__) |__ _ __ | |_   | |
   |  ___/|  ___/ __| '_ \| | |  | |
   | |    | |   \__ \ |_) | | |  | |
   |_|    |_|   |___/ .__/|_|_|  |_|
                    | |
                    |_|  by Massimo Rimondini
```
PPspliT is a PowerPoint add-in that transforms each slide of a presentation into
a sequence of slides, each displaying the contents of the original slide as they
would appear at every intermediate animation step. As such, its most natural
context of application is to produce a redistributable version of a presentation
in a _flat_ file format like PDF.

To some extent, PowerPoint already provides export functions that are meant to
include animations in the target file (e.g., it can export a presentation as a
video). However, to my knowledge, a true conversion of existing slides into an
equivalent sequence of static (i.e., animation-less) slides that is suitable for
printing or PDF export has never been natively offered by PowerPoint. PPspliT
tries to fill this gap.

----
* [Features](#features)
* [Usage](#usage)
* [Building](#building)
  * [Prerequisites](#prerequisites)
  * [Packaging for Windows](#packaging-for-windows)
  * [Packaging for MacOS](#packaging-for-macos)
* [Known limitations](#known-limitations)
* [Manual installation](#manual-installation)
* [References](#references)
* [Acknowledgments](#acknowledgments)
* [Troubleshooting](#troubleshooting)
----

## Features
* **User experience**
  * Fully integrated with PowerPoint: it is natively implemented in Visual Basic
  for Applications (VBA).
  * Adds a new tab in PowerPoint's native ribbon toolbar (or dedicated toolbar
  for PowerPoint releases prior to 2007): splitting slides is a one-click task.
  * Can operate on a range of selected slides or on the whole presentation, if no
  slides are selected.
* **Capabilities**
  * Supports all entry, emphasis, exit and motion path effects applied to slide
  shapes (with some caveats, see below).
  * Supports "Rewind when done playing", "Hide on next mouse click" and
  "Auto-reverse" effect flags, as well as reversed motion paths.
  * Can split slides at every click-triggered animation effect (like it would
  happen during a slideshow) or at each and every animation effect (useful to
  preserve multiple intermediate animations that are played without any speaker
  interaction).
  * Updates [custom
  slideshows](https://support.microsoft.com/en-gb/office/create-and-present-a-custom-show-09d4d340-3c47-4125-b177-0de3be462c5d)
  by replacing each slide with those that result from splitting it. Custom
  slideshows are actually named sequences of slides from the full deck and,
  besides being used at presentation time, they can also be selected as
  predefined slide ranges when printing. Therefore, this feature can be
  exploited to conveniently export (by printing to PDF) slide subsequences that
  are predefined in the form of custom slideshows.
  * Can optionally preserve slide numbers during splitting: if slide footers
  contain text frames with dynamically computed slide numbers, these can be
  overwritten so that numbers in all the slides resulting from splitting a single
  original slide match its original slide number.
  * Operates with native PowerPoint shapes: the slides produced after the split
  are derived from the original presentation and still contain editable shapes.
  * Format-agnostic: since the final product is still a slide deck, you can
  export it to any document format for which you have a virtual printer or file
  converter installed. PDF is implicitly supported, as PowerPoint has been
  including an export function to this format for a few years now.

Some examples displaying the operation of the add-in can be found in the
[project home page][Home page].


## Usage
Simply click on the "Split animations" button of the PPspliT toolbar.
Using the appropriate checkboxes on the same toolbar, you can choose to split
slides on animation effects that are triggered by a mouse click (most common
usage) or just every animation effect (this may be especially slow). You can
also choose to preserve slide numbers during the split.

[Usage instructions](PPspliT-howto.pdf) are also available.

*Notice*: in all releases older than 2.0 the add-in makes heavy use of the
system clipboard. Therefore, it is very important that you refrain from using it
during the split and that no programs interfere with the clipboard at all.
Effective since release 2.0, this requirement has been relaxed, and the system
clipboard can be safely used while a slide deck is being split.

**Warning**: running the add-in will modify your presentation. Even though it is
generally possible to revert the changes using the undo feature (Ctrl+Z), it is
strongly advised to work on a copy of the original slide deck to avoid losing
your work by accidentally overwriting it with the split presentation.

It may take a while for the split process to complete. If you are wondering
1. why so much code and
2. why does it take so long to split animations

here are some hints:
* PowerPoint applies slideshow effects to rasterized versions of the shapes.
Instead, in PPspliT the same effects are re-implemented on the original shape
objects.
* VBA has some sparse bugs here and there, which allow limited or no access to
shape properties. I needed to work these around to my best.
* Each animation step requires creating a new slide, which is time consuming.
* For each animation step, all the shapes that are supposed to appear later on
by means of a subsequent entry effect or to have disappeared because of a
preceding exit effect must be appropriately removed.


----

## Building
As PPspliT is implemented as a VBA macro inside PowerPoint, there is no true
_build_ procedure. The source code is embedded in PowerPoint binary files that
are saved as native PowerPoint add-ins: this is also the reason why changes are
tracked in separate files (e.g., [PPspliT.bas](src/PPT12+/PPspliT.bas)). \
The only step that requires _building_ is the generation of distributable
installers.

### Prerequisites
* Windows
  * [Nullsoft Scriptable Install System (NSIS)](https://sourceforge.net/projects/nsis/)
  * [Office 2007 Custom UI editor](http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2009/08/07/7293.aspx) -- As of June 2020, the link seems broken: you may try using the [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) instead.
* MacOS
  * Script Editor (ships natively with MacOS)

### Packaging for Windows
* Edit the VBA macro inside `PPT12+\PPspliT.pptm` as needed, then prepare the
file as follows:
  * Update the release number if required (also in the about dialog box).
  * Save the file (`PPspliT.pptm`).
  * Export each module from the Visual Basic for Applications editor into
  corresponding `.bas`, `.frm` and `.frx` files.
  * Export the whole file as a PowerPoint add-in (`PPspliT.ppam`).
  * Open `PPspliT.pptm` using the Office 2007 Custom UI Editor or the Office
  RibbonX Editor, update the release number if required, and save the file.
  * Do the same for `PPspliT.ppam`.
* Apply consistent changes to file `PPT11-\PPspliT.ppt`, save it, export each
module and export the whole PPT file as a PowerPoint 97-2003 add-in (`PPspliT.ppa`).
* Edit file `ppsplit_installer.nsi` to refresh the release number if required.
* Process file `ppsplit_installer.nsi` through NSIS (usually it is enough to
right-click on the file and select "Compile NSIS script"). File `PPspliT-setup.exe`
should then be generated in the parent folder.

### Packaging for MacOS
* Apply changes to `PPspliT.pptm` and export it as PowerPoint add-in `PPspliT.ppam`
as described above for the Windows case.
* Open file `MacOS/PPspliT for MacOS/Install PPspliT.app` using Apple's Script
Editor.
* Refresh resource `PPspliT.ppam` inside the script by dragging and dropping the
updated `PPspliT.ppam` inside the Script Editor.
* Save the installer and close the Script Editor.
* Open a Terminal window and run script `MacOS/PPspliT for MacOS/build_macos_dmg.sh`
to generate file `PPspliT.dmg`.

----

## Known limitations
Yes, the list is apparently long, but please look carefully through it because
it consists mostly of corner cases.
* PPspliT does not offer any PDF conversion functions: it is not meant to. It
just processes a presentation to split animations, then it is up to your
favorite PDF generation software or PowerPoint's native PDF export function to
generate the final PDF (or whatever other document format).
* PPspliT does not preserve animation effects: the slide deck resulting from a
split accurately renders the status of the slideshow at each intermediate
animation step, but every slide is cleared of all animation effects. This means
that you cannot have "moving shapes" in your final _flat_ (PDF) document. Even
if animations were preserved in the slides, embedding them in the final document
would require advanced processing functions for every possible output document
format, which is out of the scope of PPspliT, and would lead to much less
portable documents.
* All of the add-in features are implemented for all PowerPoint versions, but
minor glitches may exist with versions prior to 2007, sometimes due to VBA
limits or bugs.
* Some functions are knowingly unsupported and may never be implemented:
1. *Slide transitions*
&mdash; Since they are meant to smoothen slide changes, they have no
persistent effects on their contents, hence no action that needs to be
rendered by PPspliT.

2. *Shape dimming after playing an effect*

3. *Most effects/actions triggered by mouse clicks on a specific shape*
&mdash; As an exception, cross-slide hyperlinks *are* supported: their targets
are updated to point to the originally meant slides even after they have
been renumbered by the split. Slide previews using the
[zoom feature](https://support.microsoft.com/en-us/office/use-zoom-for-powerpoint-to-bring-your-presentation-to-life-9d6c58cd-2125-4d29-86b1-0097c7dc47d7)
are *not* supported anyway, meaning that they may become broken after splitting.

4. *The shaking and blinking emphasis effects*
&mdash; This is due to a PowerPoint bug.

5. *Effects applied to individual shapes of composite objects (SmartArt, charts)*
&mdash; To my knowledge, the interface exposed by VBA to alter the properties of
such shapes is somewhat limited. For example, a position property like
`Selection.ShapeRange(1).SmartArt.Nodes(1).Shapes(1).Left` is read-only, and
methods like `ScaleHeight` or `Cut` affect the whole SmartArt object despite
being applied to its individual shapes. Shape groups *are* of course supported.

6. *For emphasis effects, repetition and "Until next click" duration*
&mdash; The *duration* parameter of emphasis effects normally indicates the
time that it takes to play the effect until its end. For very few effect
types, this same setting indicates the time for which the effect persists
on its target shape instead. Effects that are not persistent (i.e., they
have an established duration in seconds) are simply ignored by PPspliT. Any
other emphasis effects are assumed to last until the end of the slide (or
until a subsequent effect is applied to the same slide). This means that
emphasis effects that last "until next click" are *not* supported and are
handled in the same way as effects that last "until end of slide".
On the other hand, the *repeat* setting allows to loop the effect's action
for an established number of iterations or, alternatively, until the next mouse
click or the end of the current slide. Since effect loops don't have any
meaningful outcome on a statically rendered slide, PPspliT simply ignores
this setting and assumes that all emphasis effects are applied once (i.e.,
without loops).

7. *Accurate rendering of color effects*
&mdash; PowerPoint implements color change effects in a way that is honestly
hard for me to reverse engineer. PPspliT approximates these effects but the
final applied color may not perfectly coincide with the one natively applied
by PowerPoint.

8. *Many emphasis and motion effects that apply to a single text paragraph
instead of a whole shape*
&mdash; In general, all those effects whose rendering requires separation of
the text frame from its parent shape are unlikely to be supported.

9. *Rasterized shape scaling and non-proportional text resizing*
&mdash; PowerPoint applies any effects to rasterized versions of the shapes. As
a consequence, grow/shrink effects affect all the elements of a shape
(including, e.g., shape border thickness) and not necessarily preserve the
aspect ratio. PPspliT resizes the native shape instead, thus preserving its
components (including border thickness) and resulting in a sharper rendering,
because the native vector shapes are preserved and there is no interpolation
introduced by resizing or rotation effects. While this is generally welcome,
the final result may sometimes differ from the intended one. Most evidently,
PPspliT only supports proportional growing/shrinking of text elements:
if a grow/shrink effect occurs on a text element and is set to only affect it
vertically or horizontally, PPspliT renders it by adjusting the font size by
an amount that is a good compromise between horizontal and vertical growth/shrink,
but no "compression" or "expansion" of the text occurs.

10. *Accurate rendering of some rotation effects*
&mdash; When a slide show is played, PowerPoint rotates shapes around the
center of the visible shape body. Instead, PPspliT rotates them around the
center of the container box. Sometimes the container box may be larger than the
visible shape, resulting in a different center of rotation being applied. To
explain the difference, consider an arc, whose container box is the rectangle
(or, possibly, square) that encloses the full circle: at slideshow time
PowerPoint can rotate the arc around the center of the arc stroke itself,
whereas PPspliT would rotate it around the center of the container box: since
the latter is generally (much) larger than the visible arc, the final
impression is that the visible shape (the arc) has "wandered around".

11. *Exit/entry effects applied to shapes that are part of a slide layout are only
partially supported*
&mdash; In fact, these shapes are turned into placeholders (instead of
disappearing altogether) when one attempts to delete them. While this is not an
issue in the vast majority of cases, if such placeholders have a formatting
applied (e.g., a background color) they may stay visible even when they are not
expected to.

12. *Adjustment of slide numbers on a PPTX file that is imported into PowerPoint
<=2003 using the Microsoft Office Compatibility Pack*
&mdash; This is a very old special condition and is never expected to occur.

13. *Adjustment of (dynamic) slide numbers that appear in standard text boxes*
&mdash; Although dynamically updated slide numbers can be inserted in any text
paragraph, PPspliT is only able to adjust them (i.e., preserve a numbering that
is coherent with the one of the original slides even after splitting) if such
numbers appear in special placeholder boxes defined in slide masters and
inserted as headers/footers in the slide deck.

14. *Animations in slide masters*

15. *Animation effects whose order is strictly dependent on timing*
&mdash; Animation effects can be played after a mouse click ("on click"), after
the preceding effect has ended ("after previous") or at the same time as a
preceding effect ("with previous"). While ordering of the effect outcomes is
strictly defined in the first two cases, it may depend on timing in the third
case. For example, if effect B follows effect A in the animation sequence, both
effects are set to play "with previous" but effect A has a *delay* set to 1
second whereas effect B has no delay, effect B is played before effect A.
PPspliT does not consider this kind of reordering, and assumes that effects are
always played in the same order in which they appear in the animation sequence.

16. *Something else I am not aware of*

---

## Manual Installation
PPspliT comes packaged with an installer which eases setup by means of a
convenient wizard. The installer is tuned to operate in many different typical
configurations (as a sole exception, security features in recent releases of
macOS may require following specific [(un)installation
instructions](https://www.maxonthenet.altervista.org/downloads/PPspliT/PPspliT-macOS-howto.pdf)).
However, in the unlikely case in which it fails, you can still attempt a
manual installation by following the procedure below.


1. Download the latest [PPspliT macro file](https://github.com/maxonthegit/PPspliT/raw/master/src/PPT12%2B/PPspliT.ppam)
and save it to a location of your choice.
As an alternative, download the latest [Windows installer file][Home page]
(even if you are using macOS), open it with your favorite Zip unpacker (do _not_ start the installer),
find file `PPspliT.ppam` inside the archive and save it to a location of your choice.



2. Start PowerPoint.
3. Add the downloaded `PPspliT.ppam` as a PowerPoint add in as described at	 https://support.microsoft.com/en-us/office/add-or-load-a-powerpoint-add-in-3de8bbc2-2481-457a-8841-7334cd5b455f,
namely:

* In case you are using PowerPoint for Windows:

4. Click the _File_ tab, then _Options_.
5. In the _Options_ dialog box, click _Add-Ins_.
6. In the _Manage_ list at the bottom of the dialog box, click _PowerPoint Add-ins_,
then click _Go_.
7. In the _Add-Ins_ dialog box, click _Add New_.
8. In the _Add New PowerPoint Add-In_ dialog box, browse for the previously saved
`PPspliT.ppam` file and then click _OK_.
9. A security notice might appear. In this case, click on _Enable Macros_ and
then click _Close_.

* If, instead, you are using PowerPoint for macOS:

4. Open the _Tools_ menu on the top bar (i.e. not the PowerPoint ribbon toolbar,
but rather the macOS menu bar at the top of the screen) and select _PowerPoint add ins_.
5. Click on _+_ and select file `PPspliT.ppam` that you saved above.

Note that,even when this issue occurs, the installer should have already
taken care of copying the required files to a proper location in your system
(usually `%APPDATA%\Microsoft\AddIns\PPspliT` for Windows systems, and
`$HOME/Library/Group\ Containers/UBF8T346G9.Office` for macOS systems).
Therefore, the steps for unpacking and saving file `PPspliT.ppam` described above
can be skipped.


---

## References
* [Project home page][Home page]
* At the time of first sketching the add-in code, I used [this blog post by Neil
Mitchell](http://neilmitchell.blogspot.com/2007/11/creating-pdf-from-powerpoint-with.html)
and [its follow-up](https://neilmitchell.blogspot.com/2007/11/powerpoint-pdf-part-2.html) as greatly inspirational starting points.

## Acknowledgments
Although I am the only developer of the add-in, several suggestions for
improvements and bug fixes came in the form of feedback from its end users. Some
of them are acknowledged in the [changelog](CHANGES.txt).

----

# Troubleshooting

_The add-in is splitting only the first slide instead of the whole slide deck._

Maybe you have accidentally selected the first slide in the left-side thumbnail
pane of PowerPoint. Just try clicking anywhere in the main pane of
PowerPoint (i.e., the slide editor) and try PPspliTting again.

---

_The (Windows) installer fails to recognize any PowerPoint releases (error
message "the add in has been left unconfigured")._

This may happen, for example, when a pre installed OEM PowerPoint release is
being used. There is not much that can be done to address this issue, as
specially packed PowerPoint releases may be harder to detect for the installer
and the effort to improve its recognition capabilities exceeds the user
reported impact of this problem. However, you can still work around this little
problem by performing a [manual installation](#manual-installation).

---

_Error "Macro cannot be found or has been disabled because of security" is
displayed every time a split is attempted._

As an outdated but, possibly, still valid explanation, a [security update
released by Microsoft](http://support.microsoft.com/kb/2598041/en-us") around
April 2012 may cause this issue with most VBA-based applications that make use
of dialog boxes, including PPspliT. To correct this problem, Microsoft suggests
deleting cached versions of control type libraries, which is harmless for your
system. I can confirm that this solution has worked for me. Basically, you have
to delete all `.exd` files stored in `%HOMEPATH%\Application
Data\Microsoft\Forms` and `%TEMP%\VBE`. Please rely on the official instructions
from Microsoft, which can be found in the page mentioned above. \
If this does not solve your problem, then either you are still using a really
outdated PPspliT release (1.5 was known to have such compatibility problems) or
your macro security settings may need to be reviewed.




[Home page]: http://www.maxonthenet.altervista.org/ppsplit.php
