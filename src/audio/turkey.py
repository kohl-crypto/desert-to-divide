import numpy as np, wave
from scipy.signal import butter, lfilter
SR=22050
def bp(x, lo, hi):
    b,a=butter(2,[lo/(SR/2),hi/(SR/2)],btype='band'); return lfilter(b,a,x)
def gobble(dur=0.85, f0=620, seed=0):
    rng=np.random.default_rng(seed)
    t=np.arange(int(SR*dur))/SR
    # descending pitch with fast warble (the "obble-obble")
    f=f0*(1-0.35*t/dur)*(1+0.08*np.sin(2*np.pi*14*t))
    ph=2*np.pi*np.cumsum(f)/SR
    src=sum((1/(k**1.1))*np.sin(k*ph) for k in range(1,14))
    src+=0.15*rng.standard_normal(len(t))   # breathiness
    # syllable gating ~ 13 per second, sharp attacks
    rate=13+rng.uniform(-1,1)
    g=np.clip(np.sin(2*np.pi*rate*t),0,1)**0.6
    env=np.minimum(1,t/0.02)*np.exp(-1.2*t/dur)
    y=bp(src,500,1400)+0.7*bp(src,1700,3200)+0.3*bp(src,3400,4800)
    y*=g*env
    return y/np.max(np.abs(y))
def read(p):
    w=wave.open(p); x=np.frombuffer(w.readframes(w.getnframes()),dtype=np.int16).astype(float)/32768; return x
voice=read('voice_hi.wav')
out=np.zeros(int(SR*(len(voice)/SR+2.2)))
g1=gobble(0.9,640,1); out[:len(g1)]+=0.8*g1
off=int(SR*0.85); out[off:off+len(voice)]+=1.0*voice
# little turkey trills peppered under the chant
for i,s in enumerate(np.linspace(1.0, 0.85+len(voice)/SR-0.4, 5)):
    g=gobble(0.35,700+30*i,10+i); k=int(SR*s); out[k:k+len(g)]+=0.25*g
g2=gobble(1.1,600,7); k=off+len(voice)-int(0.1*SR); out[k:k+len(g2)]+=0.9*g2
out=out[:k+len(g2)]
out/=np.max(np.abs(out))*1.05
w=wave.open('gobble.wav','wb'); w.setnchannels(1); w.setsampwidth(2); w.setframerate(SR); w.writeframes((out*32767).astype(np.int16).tobytes()); w.close()
print(len(out)/SR)
