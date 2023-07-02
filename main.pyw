from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from datetime import datetime
from time import sleep
from winotify import Notification, audio
from os import getcwd
from random import randint
while True:
    devices = AudioUtilities.GetMicrophone()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = interface.QueryInterface(IAudioEndpointVolume)
    multiplier = 100
    before = round(volume.GetMasterVolumeLevelScalar() * multiplier)
    if volume.GetMasterVolumeLevelScalar() > 0.81:
        print(datetime.now())
        print(before, " Громкость микрофона до")
        print(volume.GetVolumeRange(), "Радиус громкости микрофона (dB)")
        volume.SetMasterVolumeLevel(18, None) # громкость микрофона в децибелах
        after = round(volume.GetMasterVolumeLevelScalar() * multiplier)
        print(f"Понизил громокость микрофона с {before} до {after}")
        print(round(volume.GetMasterVolumeLevelScalar() * multiplier), "Громкость микрофона после")
        toast = Notification(app_id=f"Понижение громкости микрофона{randint(0, 100)}",
                             title="Автоматическая регулировка громкости микрофона",
                             msg=f"{before} -> {after}",
                             duration="short",
                             icon=getcwd() + "\pyc.ico")
        toast.set_audio(audio.Mail, loop=False)
        toast.show()
    elif volume.GetMasterVolumeLevelScalar() < 0.79:
        print(datetime.now())
        print(before, " Громкость микрофона до")
        print(volume.GetVolumeRange(), "Радиус громкости микрофона (dB)")
        volume.SetMasterVolumeLevel(18, None) # громкость микрофона в децибелах
        after = round(volume.GetMasterVolumeLevelScalar() * multiplier)
        print(f"Повысил громокость микрофона с {before} до {after}")
        print(round(volume.GetMasterVolumeLevelScalar() * multiplier), "Громкость микрофона после")
        toast1 = Notification(app_id=f"Повышение громкости микрофона{randint(0, 100)}",
                              title="Автоматическая регулировка громкости микрофона",
                              msg=f"{before} -> {after}",
                              duration="short",
                              icon=getcwd() + "\pyc.ico")
        toast1.set_audio(audio.Mail, loop=False)
        toast1.show()
    sleep(1)
