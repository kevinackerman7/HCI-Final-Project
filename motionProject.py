#CSC 355 Human Computer Interaction Final Project 
#Ben Delany, Andrew Yoon, Cole Ding, Kevin Ackerman  

#Music Interaction and Gesture Identification (MIGI)

'''
This project was made using Python 2.7 and the Leap Motion controller. The purpose of this project is for users to be able to control
music apps using the Leap Motion controller isntead of using the keybaord buttons or the in-app controls. We achieved this through using
the Windows 32 API that allowed us to bind the actions of a keyboard press to other actions, such as a gesture from the Leap Motion.
The final result of this project includes functionalities such as opening and closing the Spotify app, raising/lowering volume,
next/previous song, mute/unmute and pause/play. 
'''

import sys
import thread
import time
import win32api
import Leap
import math
import os
import Tkinter

from Tkinter import *

from Leap import KeyTapGesture, SwipeGesture, ScreenTapGesture #load in pre-existing Leap Motion gestures 

# Assign key codes for Windows to variables
VK_ESCAPE = 0x1B
VK_MEDIA_PLAY_PAUSE = 0xB3
VK_MEDIA_NEXT_TRACK = 0xB0
VK_MEDIA_PREV_TRACK = 0xB1
VK_VOLUME_MUTE = 0xAD
VK_VOLUME_DOWN = 0xAE
VK_VOLUME_UP = 0xAF

#Assign virtual keys to variables
code = win32api.MapVirtualKey(VK_ESCAPE, 0)
code1 = win32api.MapVirtualKey(VK_MEDIA_PLAY_PAUSE, 0)
code2 = win32api.MapVirtualKey(VK_MEDIA_NEXT_TRACK, 0)
code3 = win32api.MapVirtualKey(VK_MEDIA_PREV_TRACK, 0)
code4 = win32api.MapVirtualKey(VK_VOLUME_MUTE, 0)
code5 = win32api.MapVirtualKey(VK_VOLUME_DOWN, 0)
code6 = win32api.MapVirtualKey(VK_VOLUME_UP, 0)


class LeapEventListener(Leap.Listener):


    # Printed on startup of execution
    def on_init(self, controller):
        print "Initialized"
        self.prevTime = time.time()
        self.minTime = 1
        
    # Upon connection, enable desired gestures to run in the background
    def on_connect(self, controller):
        print "Connected"
        
        #enabling each pre-existing gesture so that they can be used in the recognition process
        controller.enable_gesture(Leap.Gesture.TYPE_SWIPE);
        controller.enable_gesture(Leap.Gesture.TYPE_KEY_TAP);
        controller.enable_gesture(Leap.Gesture.TYPE_SCREEN_TAP);
        controller.set_policy(Leap.Controller.POLICY_BACKGROUND_FRAMES); #setting policy flag so that the progam can run in the background
        controller.config.save(); #save configuration

    def on_disconnect(self, controller):
        print "Motion Sensor Disconnected"

    def on_exit(self, controller):
        print "Exited"

    # Handle motions while looping through frames
    def on_frame(self, controller):

        frame = controller.frame()

        #hand = frame.hands[0]

        # Open Spotify/Close
        if len(frame.hands) == 2 and (time.time() - self.prevTime) > self.minTime: #if 2 hands are in the frame and a time of 1 second has passed since the last gesture then
            self.prevTime = time.time() #reset the previous time
            if(math.fabs(frame.hands[0].grab_strength - frame.hands[1].grab_strength) > 0.95): #if one hand is making a fist and the other hand is an open palm
                print "Close Spotify"
                os.system("TASKKILL /F /IM Spotify.exe") #kill the task
            else: #if both hands are open palms
                print "Open Spotify"
                os.system("start Spotify.lnk") 

        for gesture in frame.gestures(): 

            if gesture.type == Leap.Gesture.TYPE_SWIPE: #if the detected gesture is a swipe of any type
                swipe = SwipeGesture(gesture) 
                direction = swipe.direction #save x,y,z direction values of the swipe 

                if (direction.x > 0 and math.fabs(direction.x) > math.fabs(direction.y) and (time.time() - self.prevTime) > self.minTime):
                #if the swipe is the right-> direction, and the swipe is more horizontal than it is vertical, and a time has passed since the last gesture then
                    win32api.keybd_event(VK_MEDIA_NEXT_TRACK, code2) #press the keyboard button for next track
                    self.prevTime = time.time() #update time
                    print "Next Track, and the speed is %f" % (swipe.speed)

                elif (direction.x < 0 and math.fabs(direction.x) > math.fabs(direction.y) and (time.time() - self.prevTime) > self.minTime):
                #if the swipe is the <-left direction, and the swipe is more horizontal than it is vertical, and a time has passed since the last gesture then
                    win32api.keybd_event(VK_MEDIA_PREV_TRACK, code3) #press keyboard putton for previous track
                    self.prevTime = time.time() #update time
                    print "Previous Track"

                elif (direction.y < 0 and math.fabs(direction.y) > math.fabs(direction.x)):
                #if the swipe is in the downward direction, and the swipe is more vertical than it is horizontal, then 
                    win32api.keybd_event(VK_VOLUME_DOWN, code5) #press keyboard button for volume down
                    print "Volume Down"

                elif (direction.y > 0 and math.fabs(direction.y) > math.fabs(direction.x)):
                #if the swipe is in the upward direction, and it is more vertical than it is horizontal, then
                    win32api.keybd_event(VK_VOLUME_UP, code6) #press keyboard button for volume up
                    print "Volume Up"

 
            elif gesture.type == Leap.Gesture.TYPE_KEY_TAP and (time.time() - self.prevTime) > self.minTime:
            #if the type of gesture detected is the key tap (downward tapping motion) and a time has pased since the last gesture, then
                tap = KeyTapGesture(gesture)
                win32api.keybd_event(VK_MEDIA_PLAY_PAUSE, code1) #press keyboard button for play/pause
                self.prevTime = time.time() #update time
                print "Play/Pause Track"

            # Screen tap to mute/unmute
            elif gesture.type == Leap.Gesture.TYPE_SCREEN_TAP and (time.time() - self.prevTime) > self.minTime:
            #if the gesture detected is a scree tap gesture (forward tapping motion), and a time has passed since the last gesture, then
                tapping = ScreenTapGesture(gesture)
                win32api.keybd_event(VK_VOLUME_MUTE, code4) #press keyboard button for mute/unmute
                self.prevTime = time.time() #update time
                print "Mute/Unmute"


def main():


    listener = LeapEventListener() #initialize listener 

    controller = Leap.Controller() #initialize controller

    controller.add_listener(listener) #add listener to controller 

#GUI code 
    
def openSpotify():
	print("Open Spotify")
	os.system("start Spotify.lnk")
	
def closeSpotify():
	print("Close Spotify")
	os.system("TASKKILL /F /IM Spotify.exe")
    
window = Tk()
scrollbar = Scrollbar(window)
window.geometry("650x600")
window.title("M.I.G.I.")
scrollbar.pack(side = RIGHT, fill = Y)
spc = Label(window, text = " ")
lbl2 = Label(window, text = "M.I.G.I.: Music Interaction and Gesture Identification", font = ("Arial Bold", 12), anchor = "w")
desc = Label(window, text = "	MIGI uses Leap Motion gesture identification and the Windows API to ", anchor = "w")
desc2 = Label(window, text = "	give users an easy, intuitive way to interact with their media player.", anchor = "w") 
lbl1 = Label(window, text = "Gesture Interaction Options", font = ("Arial Bold", 12), anchor = "w")
nextSong = Label(window, text = "Next Song", font = ("Arial Bold", 9), anchor = "w")
previous = Label(window, text = "Previous Song", font = ("Arial Bold", 9), anchor = "w")
play = Label(window, text = "Play/Pause", font = ("Arial Bold", 9), anchor = "w")
mute = Label(window, text = "Mute/Unmute", font = ("Arial Bold", 9), anchor = "w")
open = Button(window, text = "Open Spotify", font = ("Arial Bold", 9), anchor = "w", command = openSpotify)
close = Button(window, text = "Close Spotify", font = ("Arial Bold", 9), anchor = "w", command = closeSpotify)
volup = Label(window, text = "Raise Volume", font = ("Arial Bold", 9), anchor = "w")
voldown = Label(window, text = "Lower Volume", font = ("Arial Bold", 9), anchor = "w")

odesc = Label(window, text =  "	To open Spotify, place two open hands over the Leap Motion.", font = ("Arial", 9))
cdesc = Label(window, text =  "	To close Spotify, place one open hand and one fist over the Leap Motion.", font = ("Arial", 9))
nsdesc = Label(window, text = "	To skip to the next song, wave your open hand from left to right above the Leap Motion.", font = ("Arial", 9))
psdesc = Label(window, text = "	To skip to the previous song, wave your open hand from right to left above the Leap Motion.", font = ("Arial", 9))
ppdesc = Label(window, text = "	To play or pause music, make a downwards tapping motion with your pointer finger.", font = ("Arial", 9)) 
mudesc = Label(window, text = "	To mute or unmute music, make a forwards tapping motion towards your screen with your pointer finger.", font = ("Arial", 9)) 
vudesc = Label(window, text = "	To raise the volume of your music, swipe up, away from the Lepa Motion, with an open hand.", font = ("Arial", 9)) 
vddesc = Label(window, text = "	To lower the volume of your music, swipe down, towards the Leap Motion, with an open hand.", font = ("Arial", 9))

swipe = PhotoImage(file = "swipe.gif")
screen = PhotoImage(file = "screen.gif")
key = PhotoImage(file = "key.gif")
#open = PhotoImage(file = "twohands.gif")
#close = PhotoImage(file = "close.gif")
swipePic = Label(image = swipe)
screenTap = Label(image = screen)
keyTap = Label(image = key)
#openLabel = Label(image = open)
#closeLabel = Label(image = close)

lbl2.pack(anchor = "w", pady = (5, 5), padx = (30, 0))
desc.pack(anchor = "w")
desc2.pack(anchor = "w")
spc.pack(anchor = "w")
spc.pack()
spc.pack()

lbl1.pack(anchor = "w", pady = (5, 5), padx = (30, 0))
open.pack(anchor = "w", padx = (5, 5), pady = (5, 5))
odesc.pack(anchor = "w")
close.pack(anchor = "w", padx = (5, 5), pady = (5, 5))
cdesc.pack(anchor = "w")
nextSong.pack(anchor = "w", pady = (5, 5))
nsdesc.pack(anchor = "w")
previous.pack(anchor = "w", pady = (5, 5))
psdesc.pack(anchor = "w")
play.pack(anchor = "w", pady = (5, 5))
ppdesc.pack(anchor = "w")
volup.pack(anchor = "w", pady = (5, 5))
vudesc.pack(anchor = "w")
voldown.pack(anchor = "w", pady = (5, 5))
vddesc.pack(anchor = "w")
mute.pack(anchor = "w", pady = (5, 5))
mudesc.pack(anchor = "w")

window.mainloop()

    print "Press Enter to quit"

    try:
        sys.stdin.readline()
    except KeyboardInterrupt:
        pass
    finally:
        controller.remove_listener(listener)

if __name__ == "__main__":
    main()
