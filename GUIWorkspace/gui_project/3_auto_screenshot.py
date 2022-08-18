import time
from PIL import ImageGrab
import os

time.sleep(5) # 5초 대기 : 사용자가 준비하는 시간

for i in range(1, 31): # 2초 간격으로 10개 이미지 저장
    img = ImageGrab.grab() # 현재 스크린 이미지를 가져옴
    os.mkdir('/img')
    img.save("/img/image{}.png".format(i),True) # 파일로 저장 (image1.png ~ image10.png)
    time.sleep(10) # 2초 단위