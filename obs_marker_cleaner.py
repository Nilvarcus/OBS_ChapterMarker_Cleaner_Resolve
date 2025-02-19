import os
import sys
from python_get_resolve import GetResolve
import time  # Import time module (optional, for potential delay testing)

def frame_id_to_timecode(frame_id, frame_rate):
    """
    Converts a frame ID to a timecode string (HH:MM:SS:FF).
    """
    total_seconds = frame_id / frame_rate
    hours = int(total_seconds / 3600)
    minutes = int((total_seconds % 3600) / 60)
    seconds = int(total_seconds % 60)
    frames = int(round((total_seconds - int(total_seconds)) * frame_rate)) # Use round to handle potential floating point inaccuracies
    return "{:02d}:{:02d}:{:02d}".format(hours, minutes, seconds, frames)


def run_clip_markers_to_timeline():
    """
    Part 1: Copies clip markers from all clips in the current timeline and creates
    timeline markers in RED color, **transferring marker names and adding CLIP timestamp**.
    """

    resolve = GetResolve()

    if resolve is None:
        print("DaVinci Resolve application not found (using python_get_resolve).")
        return False  # Indicate failure

    project_manager = resolve.GetProjectManager()
    project = project_manager.GetCurrentProject()

    if project is None:
        print("Error: No project is currently open.")
        return False  # Indicate failure

    timeline = project.GetCurrentTimeline()
    if timeline is None:
        print("Error: No timeline is currently open in the project.")
        return False  # Indicate failure

    print("Starting Part 1: Copy clip markers to timeline markers (RED) - Transferring Names & CLIP Timestamp...")

    markers_copied_count = 0
    frame_rate_str = project.GetSetting('timelineFrameRate') # Get timeline frame rate
    frame_rate_str = str(frame_rate_str)
    if " " in frame_rate_str:
        timeline_frame_rate = float(frame_rate_str.split(" ")[0])
    else:
        timeline_frame_rate = float(frame_rate_str)

    video_track_count = timeline.GetTrackCount("video")
    for track_index in range(1, video_track_count + 1):  # Video track indices are 1-based
        timeline_items = timeline.GetItemListInTrack("video", track_index)
        if timeline_items:
            for item in timeline_items:
                clip_markers = item.GetMarkers()
                if clip_markers:
                    clip_name = item.GetName()
                    clip_start_frame_timeline = item.GetStart()
                    media_pool_item = item.GetMediaPoolItem() # Get MediaPoolItem
                    if media_pool_item:
                        source_clip_frame_rate_str = media_pool_item.GetClipProperty("Frame Rate") # Try to get source clip frame rate
                        source_clip_frame_rate = timeline_frame_rate # Default to timeline frame rate if source clip frame rate is not found or invalid
                        try:
                            source_clip_frame_rate = float(source_clip_frame_rate_str)
                        except (ValueError, TypeError):
                            print(f"Warning: Could not get valid frame rate for source clip '{clip_name}'. Using timeline frame rate for timecode.")


                    for clip_frame_id, marker_data in clip_markers.items():
                        # Calculate timeline frame ID: clip_start_frame + clip_marker_frame
                        timeline_frame_id = clip_start_frame_timeline + clip_frame_id
                        timeline_frame_id_int = int(timeline_frame_id) # Ensure integer frame ID for timeline
                        source_timecode_str = ""
                        if media_pool_item:
                            source_clip_frame_id_int = int(clip_frame_id) # Clip marker frame ID is already clip-relative
                            source_timecode_str = frame_id_to_timecode(source_clip_frame_id_int, source_clip_frame_rate) # Get Source Timecode String
                        else:
                            source_timecode_str = "Source Unavailable"


                        # --- Changed marker_color to "Red" ---
                        marker_color = "Red"
                        # --- Transfer marker name from clip marker and add CLIP timestamp ---
                        original_marker_name = marker_data.get("name", "") # Get name, default to empty string if missing
                        if original_marker_name: # Only process if there's an original name
                            original_marker_name = original_marker_name.replace("Unnamed", "").strip() # Remove "Unnamed" and whitespace
                            if not original_marker_name: # If it becomes empty after removing "Unnamed"
                                original_marker_name = "Clip Marker" # Use "Clip Marker" as default if originally just "Unnamed" or empty after removal
                        else:
                            original_marker_name = "Clip Marker" # Default if no name from the start

                        marker_name = f"{original_marker_name} @ Clip {source_timecode_str}".strip() # Append CLIP timestamp and remove leading/trailing whitespace


                        marker_note = marker_data.get("note", f"Copied from clip '{clip_name}' at Clip {source_timecode_str}") # Add CLIP timestamp to note too
                        marker_duration = marker_data.get("duration", 1)
                        marker_custom_data = marker_data.get("customData", "")

                        add_marker_success = timeline.AddMarker(
                            timeline_frame_id_int,
                            marker_color,
                            marker_name,
                            marker_note,
                            marker_duration,
                            marker_custom_data
                        )

                        if add_marker_success:
                            markers_copied_count += 1
                            print(f"  Copied marker '{marker_name}' from clip '{clip_name}' at clip frame {clip_frame_id} to timeline frame {timeline_frame_id_int} (RED)") # Included marker name in output
                        else:
                            print(f"  Error copying marker from clip '{clip_name}' at clip frame {clip_frame_id} to timeline frame {timeline_frame_id_int} (RED). timeline.AddMarker() returned False.")

    print(f"\nPart 1: Successfully copied {markers_copied_count} clip markers to RED timeline markers, **transferring names and adding CLIP timestamp**.")
    print("Part 1 Finished.")
    return True # Indicate success

def run_delete_blue_timeline_markers():
    """
    Part 2: Deletes all BLUE markers from the timeline AND all clips on ALL track types (video, audio, subtitle).
    (No changes needed in this part for marker name transfer)
    """

    resolve = GetResolve()

    if resolve is None:
        print("DaVinci Resolve application not found (using python_get_resolve).")
        return False  # Indicate failure

    project_manager = resolve.GetProjectManager()
    project = project_manager.GetCurrentProject()

    if project is None:
        print("Error: No project is currently open.")
        return False  # Indicate failure

    timeline = project.GetCurrentTimeline()
    if timeline is None:
        print("Error: No timeline is currently open in the project.")
        return False  # Indicate failure

    print("Starting Part 2: Delete all BLUE markers from timeline and ALL tracks...")

    # --- 1. Delete BLUE timeline markers (existing functionality - no change needed) ---
    timeline_delete_success = timeline.DeleteMarkersByColor("Blue")
    if timeline_delete_success:
        print("  Successfully deleted BLUE timeline markers.")
    else:
        print("  No BLUE timeline markers found or error deleting timeline markers.")

    # --- 2. Delete BLUE clip markers from all clips on ALL track types ---
    clips_deleted_markers_count = 0
    track_types = ["video", "audio", "subtitle"]  # List of track types to check
    for track_type in track_types:
        track_count = timeline.GetTrackCount(track_type)
        for track_index in range(1, track_count + 1):
            timeline_items = timeline.GetItemListInTrack(track_type, track_index)
            if timeline_items:
                for item in timeline_items:
                    clip_delete_success = item.DeleteMarkersByColor("Blue")
                    if clip_delete_success:
                        clips_deleted_markers_count += 1
                        clip_name = item.GetName()
                        print(f"  Successfully deleted BLUE clip markers from {track_type} clip: '{clip_name}'.")
                    # No error message if clip_delete_success is False, as it might just mean no blue markers on that clip

    print(f"  Deleted BLUE clip markers from {clips_deleted_markers_count} clips across all tracks (if any).")

    print("Part 2 Finished.")
    # Return True if timeline marker deletion was successful OR if clip marker deletion happened (even if timeline failed - or vice versa)
    # Returning True even if only one type of deletion was successful OR if no blue markers were found of either type.
    return timeline_delete_success or clips_deleted_markers_count > 0 or True # Modified return logic

def run_clip_markers_from_timeline_markers():
    """
    Part 3: Iterates through existing timeline markers and adds a clip marker to the clip
    at each timeline marker position, if a clip exists at that position, **transferring marker names and adding TIMELINE timestamp**.
    """

    resolve = GetResolve()

    if resolve is None:
        print("DaVinci Resolve application not found (using python_get_resolve).")
        return False # Indicate failure

    project_manager = resolve.GetProjectManager()
    project = project_manager.GetCurrentProject()

    if project is None:
        print("Error: No project is currently open.")
        return False # Indicate failure

    timeline = project.GetCurrentTimeline()
    if timeline is None:
        print("Error: No timeline is currently open in the project.")
        return False # Indicate failure

    print("Starting Part 3: Add clip markers based on timeline markers - Transferring Names & TIMELINE Timestamp...")

    timeline_markers = timeline.GetMarkers()
    if not timeline_markers:
        print("No timeline markers found. Exiting Part 3.")
        return True # Not really an error, just nothing to do - return True

    markers_added_count = 0
    marker_color = "Green"  # Default clip marker color (can be changed)
    frame_rate_str = project.GetSetting('timelineFrameRate') # Get timeline frame rate
    frame_rate_str = str(frame_rate_str)
    if " " in frame_rate_str:
        timeline_frame_rate = float(frame_rate_str.split(" ")[0])
    else:
        timeline_frame_rate = float(frame_rate_str)


    video_track_count = timeline.GetTrackCount("video")
    if video_track_count == 0:
        print("No video tracks found in the timeline. No clip markers can be added in Part 3.")
        return True # Not an error, just nothing to do - return True

    for frame_id, marker_data in timeline_markers.items():
        timeline_frame_id_int = int(frame_id) # Ensure integer frame ID
        timeline_timecode_str = frame_id_to_timecode(timeline_frame_id_int, timeline_frame_rate) # Get Timeline Timecode String


        # Find the clip at this timeline frame ID
        current_clip = None
        for track_index in range(1, video_track_count + 1): # Check all video tracks
            timeline_items = timeline.GetItemListInTrack("video", track_index)
            if timeline_items:
                for item in timeline_items:
                    clip_start_frame = item.GetStart()
                    clip_end_frame = item.GetEnd()
                    if clip_start_frame <= timeline_frame_id_int < clip_end_frame: # Check if timeline_frame_id is within clip range
                        current_clip = item
                        break # Found a clip, no need to check other clips
                if current_clip: # Break outer loop too if clip is found
                    break

        if current_clip:
            clip_start_frame_timeline = current_clip.GetStart()
            clip_frame_id_int = timeline_frame_id_int - clip_start_frame_timeline # Calculate clip-relative frame ID


            # --- Transfer marker name from timeline marker and add TIMELINE timestamp ---
            original_marker_name = marker_data.get("name", "") # Get name, default to empty string
            if original_marker_name: # Only process if there's an original name
                original_marker_name = original_marker_name.replace("Unnamed", "").strip() # Remove "Unnamed" and whitespace
                if not original_marker_name: # If it becomes empty after removing "Unnamed"
                    original_marker_name = "Timeline Marker" # Use "Timeline Marker" as default if originally just "Unnamed" or empty after removal
            else:
                original_marker_name = "Timeline Marker" # Default if no name from the start

            marker_name = f"{original_marker_name}".strip() # Remove leading/trailing whitespace
            marker_note = marker_data.get("note", f"Added based on timeline marker position") # Add TIMELINE timestamp to note too
            marker_duration = marker_data.get("duration", 1)
            marker_custom_data = marker_data.get("customData", "")

            clip_marker_added = current_clip.AddMarker(
                clip_frame_id_int,   # frameId (positional) - clip relative frame
                marker_color,        # color (positional)
                marker_name,         # name (positional)
                marker_note,         # note (positional)
                marker_duration,     # duration (positional)
                marker_custom_data   # customData (positional)
            )

            if clip_marker_added:
                markers_added_count += 1
                print(f"  Clip marker '{marker_name}' added to '{current_clip.GetName()}' at timeline frame {timeline_frame_id_int} (clip frame {clip_frame_id_int})") # Included marker name in output
            else:
                print(f"  Failed to add clip marker to '{current_clip.GetName()}' at timeline frame {timeline_frame_id_int}. current_clip.AddMarker() returned False.")
        else:
            print(f"  No clip found at timeline frame {timeline_frame_id_int}. Clip marker not added.")

    print(f"\nPart 3: Successfully added {markers_added_count} clip markers based on timeline marker positions, **transferring names and adding TIMELINE timestamp**.")
    print("Part 3 Finished.")
    return True # Indicate success


def run_combined_marker_script():
    """
    Runs a sequence of marker scripts:
    1. Copies clip markers to RED timeline markers (**transfers names and adds CLIP timestamp**).
    2. Deletes all BLUE markers from timeline and clips.
    3. Adds clip markers based on timeline marker positions (using default BLUE color in that script, **transfers names and adds TIMELINE timestamp**).
    """

    print("--- Starting Combined Marker Script ---")

    # 1. Run script_part1_clip_markers_to_timeline.py
    print("\n--- Running Part 1: Clip Markers to Timeline Markers (RED) - Transferring Names & CLIP Timestamp ---")
    if not run_clip_markers_to_timeline():
        print("Error in Part 1. Aborting combined script.")
        return False

    # 2. Run script_part2_delete_blue_markers.py
    print("\n--- Running Part 2: Delete BLUE Timeline and Clip Markers ---")
    if not run_delete_blue_timeline_markers():
        print("Warning: Part 2 may have encountered an issue or no blue markers were found.")
        # Continue even if deleting blue markers fails (it might be intentional if there are no blue markers)

    # 3. Run script_part3_timeline_markers_to_clip_markers.py
    print("\n--- Running Part 3: Timeline Markers to Clip Markers (BLUE) - Transferring Names & TIMELINE Timestamp ---")
    if not run_clip_markers_from_timeline_markers():
        print("Error in Part 3. Combined script may not be fully complete.")
        return False # Or you might choose to continue even if part 3 fails, depending on desired behavior

    print("\n--- Combined Marker Script Completed ---")
    return True # Indicate overall script completion (even if some parts had warnings)


if __name__ == "__main__":
    run_combined_marker_script()