from __future__ import division #so that 1/3=0.333 instead of 1/3=0
from psychopy import visual, core, data, event, logging, gui, sound
from random import shuffle, choice
import numpy as np
from numpy.random import random, randint, normal, shuffle
from numpy import sin, cos, tan, log, log10, pi, average, sqrt, std, deg2rad, rad2deg, linspace, asarray
import time, os, sys, math, random#, xlwt
from game_functions import task_function, feedback
# from game_functions import task_function, practice, feedback
from practice import practice_functions
# if __name__ != '__main__': from Feedback import feedback

touchscreen = True

class Star_Game(practice_functions):

    def __init__(self, win):
        self.fn = os.path.dirname(__file__)
        self.trialClock=core.Clock()

        #file paths
        self.image_path = 'Images/Tasks/'
        self.audio_path = 'Audio/General/'
        self.aud_practice_path = 'Audio/Practice/'
        self.aud_inst_path = 'Audio/Instructions/'

        #create practice instructions
        self.practice_instructions1 = visual.TextStim(win, units='pix', pos=[0,0], height=20, text='Practice set 1: administrator demonstrates to child')
        self.practice_instructions2 = visual.TextStim(win, units='pix', pos=[0,0], height=20, text='Practice set 2: administrator walks through trials with child')
        self.practice_instructions3 = visual.TextStim(win, units='pix', pos=[0,0], height=20, text='Practice set 3: child completes trials on his/her own')
        self.practice_instructions4 = visual.TextStim(win, units='pix', pos=[0,0], height=20, text="Let's do some more practice")
        self.try_again = visual.TextStim(win, units='pix', pos=[0,0], height=20, text="Let's try that again.")

        # self.practice_cue1 = visual.TextStim(win, units=u'pix', wrapWidth=700, pos=[0,0],height=28,text="  Let's do some practice.\n\nTouch anywhere to begin.")
        # self.practice_cue2 = visual.TextStim(win, units=u'pix', wrapWidth=700, pos=[0,0],height=28,text='Touch anywhere to do some more practice.')
        # self.practice_cue3 = visual.TextStim(win, units=u'pix', wrapWidth=700, pos=[0,0],height=28,text="Are you ready to begin?")

        # #initializing audio files for practice and instructions
        # self.practice_aud1 = sound.Sound(self.aud_practice_path + 'practice_cue1.wav')
        # self.practice_aud2 = sound.Sound(self.aud_practice_path + 'practice_cue2.wav')
        # self.practice_aud3 = sound.Sound(self.aud_practice_path + 'practice_cue3.wav')
        
        #time constrains
        self.t_blank = 2
        self.t_twinkle = 2
        self.t_mask = 1
        self.timer_limit = 12

        #repeat and continue button
        self.repeat=visual.ImageStim(win=win, name='repeat_button', image= self.image_path + 'black_button.png', units=u'pix', pos=[350, -300], size=[75,75], color=[1,1,1], colorSpace=u'rgb', opacity=1.0)
        self.cont=visual.ImageStim(win=win, name='continue_button', image= self.image_path + 'black_button.png', units=u'pix', pos=[420, -300], size=[75,75], color=[1,1,1], colorSpace=u'rgb', opacity=1.0)

        #trial variables
        self.t_twinkle = 3#when the star will twinkle
        self.bintang = visual.ImageStim(win=win, name='bintang', image = self.image_path + '/star.png', units = 'pix', ori = 0, pos = [0,0], size = [60, 60], opacity = 1, mask =None, interpolate = True)#stimulus
        self.twinkle = visual.ImageStim(win=win, name='twinkle', image = self.image_path + '/twinklingstar.png', units=u'pix', ori = 0, pos = [0,0], size = [62, 62], opacity = 1, mask =None, interpolate = True)#twinkle
        self.twinkle2 = visual.ImageStim(win=win, name='twinkle2', image = self.image_path + '/twinklingstar.png', units=u'pix', ori = 0, pos = [0,0], size = [62, 62], opacity = 1, mask =None, interpolate = True)#twinklefor END routine
        self.star_selected = visual.ImageStim(win=win, name='star_selected', image = self.image_path + '/star_selected.png', units = 'pix', ori = 0, size = [60, 60])
        self.drag = visual.ImageStim(win=win, name = 'drag', image = self.image_path + '/star2.png', units = 'pix', ori = 0, pos = [0,0], size = [60, 60], opacity = 1, mask =None, interpolate = True)
        self.circledrag = visual.Circle(win, name = 'circledrag', units = u'pix', radius = 30, ori=0, pos = [0,0])
        self.circletwinkle = visual.Circle(win, name = 'circletwinkle', units = u'pix', radius = 30, ori=0, pos = [0,0])
        self.mask = visual.ImageStim(win, name='mask2', image = self.image_path + '/mask.jpg', units=u'pix', ori=0, pos=[0, 0], size=[1500,850], opacity = 1, mask =None, interpolate = True)
        self.blank=visual.TextStim(win, ori=0, font=u'Arial', pos=[0, 0], color=u'white', text='+', height=30)
        self.mouse=event.Mouse(win=win); self.mouse.getPos()


        #start feedback
        self.fb=feedback.fb(win)
        self.tf=task_function.task_functions(win)

        # number of degrees to exclude for star placement-- degrees given will be excluded on each side of the cardinal directions
        self.cardinal_exclusion_range = 7

        #create possible star positions in degrees
        self.degree_possibilities = []
        for x in [45,135,225,315]: self.degree_possibilities.extend(range(x-(44-self.cardinal_exclusion_range), x+(45-self.cardinal_exclusion_range)))

    # def run_instructions(self, win, task):
    #     self.tf.run_instruction_functions(win,task)

    def run_practice(self, win, task, grade):
        "Run practice"

        inst_set = [[self.practice_instructions1,None,None],[self.practice_instructions2,None],[self.practice_instructions3,None],[self.practice_instructions4,None]]
        aud_set = [[None,None,None],[None,None],[None,None],[None,None]]
        stim_set = [[150,100,115],[250,200],[200,150]]
        var = 'star_task'
        score_cond = [[1,0,1],[1,1],[1,1],[1,1]]
        stim_repeat = [200,150]
        
        return self.run_practice_functions(win, grade, inst_set, aud_set, stim_set, stim_repeat, score_cond, var, task)


    def run_game(self, win, grade, thisIncrement, var):
        "Run one iteration of the game with self.trialList as conditions."
        return self.run_trial(win, thisIncrement, var)

    def run_trial(self, win, thisIncrement, var):
        sz= thisIncrement
        theseKeys = ""

        r = 250
        degree = choice(self.degree_possibilities)
        radians = degree*(2*math.pi/360)
        x = r*math.cos(radians)
        y = r*math.sin(radians)
        self.bintang.setPos([x,y]); self.bintang.setSize([sz, sz])
        self.twinkle.setPos([x,y]); self.twinkle.setSize([sz, sz])
        self.twinkle2.setPos([x,y]); self.twinkle2.setSize([sz, sz])
        self.circletwinkle.setPos([x,y]); self.circletwinkle.setRadius([sz/2]); self.circletwinkle.setLineColor('#f50af2')
        self.circledrag.setPos([0,0]); self.circledrag.setRadius([sz/2]); self.circledrag.setLineColor('white')
        self.drag.setPos([0,0]); self.drag.setSize([sz, sz]); self.drag.setImage(self.image_path + '/star2.png')

        print 'degree:', degree
        print 'radians:', radians
        print 'x, y:', x, y


        #initializing variables
        score=None
        self.mouse.setVisible(0)
        first_click_time = None
        second_click_time = None
        status = 'NOT_STARTED'
        self.drag.setImage(self.image_path + '/star2.png')

        #time constrains
        t1 =  self.t_twinkle
        t2 =  self.t_twinkle + self.t_mask
        tf =  self.t_twinkle + self.t_mask + self.timer_limit

        #display fixation with repeat, pause & continue button
        task_status = self.tf.fixation_function(win)
        print '*********task_status',task_status
        
        if task_status=='repeat_task': 
            return task_status

        elif task_status=='continue_task':
            t=0; self.trialClock.reset()
            
            # present twinkling star and then put up mask
            while t<=t2:
                t=self.trialClock.getTime()
                if t<=t1:
                    self.bintang.draw()
                    self.twinkle.setOpacity((math.sin((2*math.pi)*6*(t-2)))*0.5 + 0.5)
                    self.twinkle.draw()
                if t>t1 and t<=t2: self.mask.draw()
                theseKeys = event.getKeys()
                if len(theseKeys)>0:
                    if theseKeys[-1] in ['q','escape']: return 'QUIT'
                win.flip()

            #allow participant to move star and make response, then check if correct
            self.mouse.getPos()
            start_time = self.trialClock.getTime()
            while score==None:
                if t>t2 and t<=tf:
                    t=self.trialClock.getTime()
                    self.drag.draw()
                    if self.mouse.mouseMoved() or (self.mouse.getPressed()==[1,0,0]):
                        if self.drag.contains(self.mouse.getPos()):
                            status='STARTED'
                            first_click_time = t - start_time
                            self.drag.setImage(self.image_path + '/star_selected.png')
                    if status == 'STARTED' and (self.mouse.mouseMoved() or (self.mouse.getPressed()==[1,0,0])) and t >= first_click_time + start_time + 0.5:
                        second_click_time = t - start_time
                        self.drag.setImage(self.image_path + '/star2.png')
                        self.drag.setPos(self.mouse.getPos())
                        x_resp = self.drag.pos[0]
                        y_resp = self.drag.pos[1]
                        distance = ((y_resp - y)**2 + (x_resp - x)**2)**(0.5)
                        score = int(distance<=sz)
                    if event.getKeys(keyList=['q', 'escape']):
                        return 'QUIT'
                    win.flip()
                    self.circledrag.setPos(self.drag.pos)
                    self.circletwinkle.setPos(self.twinkle2.pos)
                if t>tf:
                    # score=0
                    if not first_click_time: first_click_time=np.nan
                    score, x_resp,y_resp,distance,second_click_time = (0,np.nan,np.nan,np.nan,np.nan)

            #give feedback
            self.fb.present_fb(win,score,[self.twinkle2,self.circletwinkle,self.drag,self.circledrag])

            #write data #headers are ['Trial Number', 'Difficulty','Score','Resp Time','Adaptive']

            output = {
                'threshold_var': float(sz),
                'level': None,
                'score': int(score),
                'resp_time': float(second_click_time-first_click_time),
                'spatial_click1': float(first_click_time),
                'spatial_click2': float(second_click_time),
                'resp_pos': "(%f, %f)"%(x_resp, y_resp),
                'target_pos': "(%f, %f)"%(x,y),
                'resp_target_dist': float(distance),
            }
            
            print output
            return output
