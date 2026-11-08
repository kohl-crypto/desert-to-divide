import ctypes, sys, wave, espeakng_loader as E
lib = ctypes.CDLL(E.get_library_path())
CB = ctypes.CFUNCTYPE(ctypes.c_int, ctypes.POINTER(ctypes.c_short), ctypes.c_int, ctypes.c_void_p)
buf = []
def cb(wav, n, ev):
    if wav and n > 0: buf.extend(wav[:n])
    return 0
cbf = CB(cb)
lib.espeak_Initialize.restype = ctypes.c_int
sr = lib.espeak_Initialize(2, 0, E.get_data_path().encode(), 0)  # AUDIO_OUTPUT_SYNCHRONOUS
lib.espeak_SetSynthCallback(cbf)
lib.espeak_SetVoiceByName(sys.argv[3].encode())
# rate=1, volume=2, pitch=3, range=4
lib.espeak_SetParameter(1, int(sys.argv[4]), 0); lib.espeak_SetParameter(3, int(sys.argv[5]), 0); lib.espeak_SetParameter(4, 99, 0)
t = sys.argv[1].encode()
lib.espeak_Synth(t, len(t)+1, 0, 0, 0, 0x01, None, None)
lib.espeak_Synchronize()
w = wave.open(sys.argv[2], 'wb'); w.setnchannels(1); w.setsampwidth(2); w.setframerate(sr)
w.writeframes(ctypes.string_at((ctypes.c_short*len(buf))(*buf), len(buf)*2)); w.close()
print(sr, len(buf)/sr)
