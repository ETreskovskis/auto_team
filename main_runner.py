import argparse
import sys
from functools import partial

from auto_join_teams_meeting import OutlookApi, IUIAutomation, EnumActiveWindows, MouseEvents, TeamsRunner

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Teams AUTO-JOIN. For additional parameter info use --help")
    parser.add_argument("--mic", type=str, required=True, default="off",
                        help="Provide flag for microphone: 'on' or 'off'. Note: this set up for all upcoming meetings",
                        )
    parser.add_argument("--camera", type=str, required=True, default="off",
                        help="Provide flag for camera: 'on' or 'off'. Note: this set up for all upcoming meetings",
                        )
    parser.add_argument("--start_before", type=int, required=False,
                        help="Provide time (seconds) to join before actual meeting has started",
                        default=3 * 60)

    arguments = parser.parse_args()

    outlook_class = OutlookApi(time_before=arguments.start_before)
    planned_meetings = outlook_class.available_meetings()
    wrapp_iui_auto = partial(IUIAutomation, camera=arguments.camera_state, mic=arguments.mic_state)
    enum_class = EnumActiveWindows()
    mouse_event = MouseEvents()
    run_meetings_bool, run_meetings_list = TeamsRunner.run_meetings(planned_meetings, enum=enum_class,
                                                                    iui_auto=wrapp_iui_auto,
                                                                    outlook=outlook_class, mouse=mouse_event)
    if not run_meetings_bool:
        sys.exit("There are no meetings to start. Quiting.")
    sys.exit(f"Quiting threads. Finished meetings: {*run_meetings_list,}")