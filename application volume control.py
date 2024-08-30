from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
import math

devives = AudioUtilities.GetSpeakers()
interface = devives.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

# volume.SetMasterVolumeLevel(-12.0, None)
# currentVolumeDb = volume.GetMasterVolumeLevel()
# volume.SetMasterVolumeLevel(currentVolumeDb - 6.0, None)

sessions = AudioUtilities.GetAllSessions()

for session in sessions:
    volume = session._ctl.QueryInterface(ISimpleAudioVolume)
    if session.Process and session.Process.name() == "msedge.exe":
        volume.SetMasterVolume(0.6, None)
    # if session.Process:
    #     print(session.Process.name())
