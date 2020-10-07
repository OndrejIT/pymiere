"""
Pymiere import and transitions
"""
import os
import win32gui
import pymiere
from pymiere import wrappers


MOVIES_TO_IMPORT = [r"C:\Users\Quentin\Desktop\VID_20180721_210002.mp4",
                    r"C:\Users\Quentin\Desktop\VID_20180721_210446.mp4",
                    r"C:\Users\Quentin\Desktop\VID_20180721_210825.mp4"]


def import_files(filepaths):
    # import files in root bin
    root_bin = pymiere.objects.app.project.getInsertionBin()
    pymiere.objects.app.project.importFiles(filepaths, True, root_bin, False)
    # return imported project items
    items = list()
    for filepath in filepaths:
        result = root_bin.findItemsMatchingMediaPath(filepath, True)
        if len(result) == 0:
            raise ImportError("Failed to find the imported item {}".format(filepath))
        if len(result) != 1:
            raise ValueError("Import sucessfull but there are more than one clips matching path {} in the root bin".format(filepath))
        items.append(result[0])
    return items


def main():
    # cannot set things in premiere without having the focus
    win32gui.SetForegroundWindow(win32gui.FindWindow("Premiere Pro", None))

    # check that a project is opened
    project_opened, sequence_active = wrappers.check_active_sequence(crash=False)
    if not project_opened:
        raise ValueError("Please open a project")

    # create new sequence from preset
    print("Create new sequence...")
    sequence_preset = os.path.realpath(os.path.join(__file__, "../base.sqpreset"))
    pymiere.objects.qe.project.newSequence("mySequence", sequence_preset)
    sequence = pymiere.objects.app.project.activeSequence
    print("Sequence created\n")

    # import and insert movies into sequence
    print("Importing clips...")
    clip_items = import_files(MOVIES_TO_IMPORT)
    import_timecode = 0  # keep track of the timecode at which we want to import starting at 0
    for i, clip_item in enumerate(clip_items):
        print(clip_item.getMediaPath())
        # each clip is imported into a new track; we are piling them up, first on top
        pymiere.objects.qe.project.getActiveSequence().addTracks(1)  # 1 = video / 0 = audio
        sequence.videoTracks[0].insertClip(clip_item, wrappers.time_from_seconds(import_timecode))
        import_timecode += clip_item.getOutPoint(1).seconds
        import_timecode -= 4  # 4 second overlap between clips
    print("Clips imported\n")

    # trim last clip to 5 second (demo purpose / TODO : doesn't trim sound associated with clip)
    print("Trimming clip...")
    last_clip = sequence.videoTracks[0].clips[-1]
    last_clip.end = wrappers.time_from_seconds(last_clip.start.seconds + 5)
    print("Clip trimmed\n")

    # add transitions
    print("Adding transitions...")
    transition_type = pymiere.objects.qe.project.getVideoTransitionByName("Iris Round")  # name are the same than in UI
    for i, clip_item in enumerate(clip_items[:-1]):
        # use qe api to add transition, we have to get back our track and clip objects
        track = pymiere.objects.qe.project.getActiveSequence().getVideoTrackAt(len(clip_items) - 1 - i)
        # here we find the first video clip (not empty clip), didn't find better way to do this with qe api
        index = 0
        clip = None
        while clip is None or clip.type == "Empty":
            clip = track.getItemAt(index)
            index += 1
        print(clip.name)
        # awesome sauce
        clip.addTransition(transition_type)
        # transitions parameters are exposed via qe api like reverse (see full list using transition_qe.inspect())
        transition_qe = track.getTransitionAt(0)
        transition_qe.setReverse(True)
        # the only thing not in qe but in official api is the duration of the transition, so with have to find back the transistion object in the other api
        transition_offi = pymiere.objects.app.project.activeSequence.videoTracks[len(clip_items) - 1 - i].transitions[0]
        transition_offi.start = wrappers.time_from_seconds(transition_offi.end.seconds - 4)  # duration of 4 seconds
    print("Transitions added\n")


if __name__ == "__main__":
    main()
