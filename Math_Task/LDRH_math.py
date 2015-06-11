# -*- coding: utf-8 -*-
from psychopy import core, visual, gui, data, misc, event, sound
import time, numpy, os, sys
from random import shuffle
from game_functions import task_function, feedback
# from game_functions import task_function, practice, feedback
from practice import practice_functions

#touchscreen? if False, uses conventional mouse
touchscreen = True

#white rectangle instead of blue button
white_rectangle = False
#dark button instead of light blue button
dark_button = False

class Math_Game(practice_functions):

    def __init__(self, win, conditions):

        #file paths for importing conditions, images and audio
        self.dir = os.path.dirname(__file__)
        image_path = 'Images/Tasks/'
        audio_path = 'Audio/General/'
        aud_practice_path = 'Audio/Practice/'
        aud_inst_path = 'Audio/Instructions/'
        self.math_dotstims_path = 'Images/Stimuli/Math_dotstims/'


        ## initialize trial components ##

        #time components and time constrains for trial
        self.trialClock = core.Clock()
        self.timer_limit = 12

        #trial condition
        self.trialList=conditions

        #mouse
        self.mouse=event.Mouse(win=win)
        self.mouse.getPos()

        #start feedback
        self.fb=feedback.fb(win)
        self.tf=task_function.task_functions(win)

        #create a dictionary to keep track of how many times you've displayed each difficulty level
        self.iteration = {}
        for operation in ['addition','subtraction','multiplication','division']:
            self.iteration[operation] = {}
            for question in range(len(self.trialList[operation])):
                self.iteration[operation][question] = 0


        ## initialize text, audio & image stimuli ##

        #practice instructions texts
        self.practice_cue1 = visual.TextStim(win, units=u'pix', wrapWidth=700, pos=[0,0],height=28,text="Touch anywhere to begin.")
        self.practice_cue2 = visual.TextStim(win, units=u'pix', wrapWidth=700, pos=[0,0],height=28,text="Let's do some more.")
        self.practice_cue3 = visual.TextStim(win, units=u'pix', wrapWidth=700, pos=[0,0],height=28,text="Touch anywhere to begin.")

        #audio files for practice and instructions
        self.practice_aud1 = sound.Sound(aud_practice_path + 'practice_cue1.wav')
        self.practice_aud3 = sound.Sound(aud_practice_path + 'practice_cue3.wav')

        #create stimuli
        self.text_stimulus = visual.TextStim(win, pos=[0,200],height=80, text='Stimulus.')
        self.dot_stimulus = visual.ImageStim(win,image=None,pos=[0,180],size=[260,260])
        self.fixation = visual.ImageStim(win, color='black', image=None, mask='circle',size=5)
        
        #for stimuli texts
        self.target = visual.TextStim(win, pos=[0,0],height=70, text='Target.')
        self.foil1 = visual.TextStim(win, pos=[0,0],height=70, text='Foil1')
        self.foil2 = visual.TextStim(win, pos=[0,0],height=70, text='Foil2')
        self.foil3 = visual.TextStim(win, pos=[0,0],height=70, text='Foil3')
        
        #for 2 & 4 stimuli buttons
        self.target_2button = visual.ImageStim(win,image= image_path + '/general_button.png')
        self.foil_2button = visual.ImageStim(win, image= image_path + '/general_button.png')
        self.target_4button = visual.ImageStim(win, image= image_path + '/general_button_4.png')#, size=[300,120])
        self.foil1_4button = visual.ImageStim(win, image= image_path + '/general_button_4.png')#, size=[300,120])
        self.foil2_4button = visual.ImageStim(win, image= image_path + '/general_button_4.png')
        self.foil3_4button = visual.ImageStim(win, image= image_path + '/general_button_4.png')


    def run_practice(self, win, task, grade):
        "Run practice"

        #instruction texts
        inst_set=[self.practice_cue1,None,None,self.practice_cue2,self.practice_cue3]
        
        #instruction audio
        aud_set=[self.practice_aud1,None,None]+[None]*2
        
        #stimuli set for practice
        stim_set = [13,11,11]+[None]*2
        
        #stimuli for repeated practice (currently the same as the initial stimuli set)
        stim_repeat = stim_set
        
        #variable, needed for some task's trial
        var = 'addition'
        
        #score condition, whether we want to constrain trial to be correct or incorrect
        score_cond = [None]*5
        
        return self.run_practice_functions(win, grade, inst_set, aud_set, stim_set, stim_repeat, score_cond, var, task)

    def run_game(self, win, grade, thisIncrement, operation):
        "Run one iteration of the game with self.trialList as conditions."
        return self.run_trial(win, thisIncrement, operation)

    def run_trial(self, win, thisIncrement, operation):
        "Run one iteration of the game."
        these_conditions = self.trialList[operation]
        this_iteration = self.iteration[operation]

        index = None

        #set the index to the current difficulty level for indexing into the conditions file
        for question in range(len(these_conditions)):
            #'self.difficulty' increases in difficulty as numbers increase, thisIncrement increases in difficulty as numbers decrease
            if int(these_conditions[question]['Difficulty']) == (len(these_conditions)-thisIncrement):
                index = question

        if index == None:
            print 'could not find index for', thisIncrement, 'in', range(len(these_conditions))
            index = 0

        difficulty = these_conditions[index]['Difficulty']
        print 'Difficulty is:', difficulty

        # Ensure iteration does not exceed length of available trials:
        if this_iteration[index] > len(these_conditions[index]['Stimulus'])-1:
            this_iteration[index] = 0

        #randomize the position of the target and foil
        four_xpositions = {-360:'left', -120:'mid-left', 120:'mid-right', 360:'right'}
        two_xpositions = {-260:'left', 260:'right'}
        target_string = str(these_conditions[index]['Correct'][this_iteration[index]])
        foil1_string = str(these_conditions[index]['Foil1'][this_iteration[index]])
        
        if these_conditions[index]['Foil2'] and these_conditions[index]['Foil3']: 
            total_foil=3
            foil2_string = str(these_conditions[index]['Foil2'][this_iteration[index]])
            foil3_string = str(these_conditions[index]['Foil3'][this_iteration[index]])
            foil_string = [foil1_string,foil2_string,foil3_string]
            foil_text = [self.foil1,self.foil2,self.foil3]
            foil_button = [self.foil1_4button,self.foil2_4button,self.foil3_4button]
            target_button = self.target_4button
            pos = four_xpositions
            xpositions = four_xpositions.keys()
        else: 
            total_foil=1
            foil2_string=''
            foil3_string=''
            foil_string = [foil1_string]
            foil_text = [self.foil1]
            foil_button = [self.foil_2button]
            target_button = self.target_2button
            pos = two_xpositions
            xpositions = two_xpositions.keys()
            

        points = [1,0,0,0]
        shuffle(xpositions)
        object_var = zip(points,[target_string]+foil_string,[self.target]+foil_text,xpositions,[target_button]+foil_button)
        for pts,string,text,xpos,button in object_var:
            text.setText(string)
            text.setColor('white')
            text.setPos([xpos,-200])
            button.setPos([xpos,-200])

        #set the text for the stimuli
        stim_string = these_conditions[index]['Stimulus'][this_iteration[index]]

        #correct division signs if applicable
        if '*div' in stim_string: stim_string= stim_string.replace('*div',u'÷').replace('-',u'−')#stim_string[0:stim_string.index('*div')] + u'÷' + stim_string[stim_string.index('*div')+4:]
        #get image if applicable
        if '.png' in stim_string:
            self.dot_stimulus.setImage(self.math_dotstims_path+stim_string)
            self.stimulus = self.dot_stimulus
        else:
            self.text_stimulus.setText(stim_string)#[self.iteration[index]]))
            self.stimulus=self.text_stimulus

        #display fixation with repeat, pause & continue button
        task_status = self.tf.fixation_function(win)
        print '*********task_status',task_status
        
        if task_status in ['QUIT','repeat_task']:
            return task_status

        elif task_status=='continue_task':
            t=0; self.trialClock.reset()

            start_time = self.trialClock.getTime()
            choice_time, score, thisResp, thisResp_pos = 0, None, None, None
            double_click, double_time, double_time2, double_time3 = False, None, None, None
            self.mouse.getPos()
            
            while score==None:
                t = self.trialClock.getTime()
                if t<=self.timer_limit:
                    self.stimulus.draw()

                    for text,button in zip([self.target]+foil_text,[target_button]+foil_button):
                        button.draw()
                        text.draw()
                    win.flip()

                    while thisResp==None and choice_time<=self.timer_limit:
                        if (self.mouse.mouseMoved() or (self.mouse.getPressed()==[1,0,0])):
                            for pts,string,text,xpos,button in object_var:
                                if button.contains(self.mouse):
                                    score,thisResp,thisResp_pos = (pts,string,pos[xpos])
                                    text.setColor('gold')
                        
                        #get key inputs
                        key = event.getKeys()
                        if key==['escape'] or key==['period']*3: return 'QUIT'
                        #check for triple click
                        if double_time and not double_time2 and self.trialClock.getTime()-double_time>=1: double_click, double_time = False, None
                        elif double_time2 and not double_time3 and self.trialClock.getTime()-double_time2>=1: double_click, double_time, double_time2 = False, None, None
                        
                        if double_click==False and key==['period']:
                            double_click = 'maybe'
                            double_time = self.trialClock.getTime()
                        elif double_click=='maybe' and key==['period']:
                            double_time2 = self.trialClock.getTime()
                            if double_time2 - double_time >1: double_click, double_time, double_time2 = False, None, None
                            elif double_time2 - double_time<=1:double_click = 'yes'
                        elif double_click=='yes' and key==['period']:
                            double_time3 = self.trialClock.getTime()
                            if double_time3-double_time2>1: double_click, double_time, double_time2, double_time3 = False, None, None, None
                            elif double_time3-double_time2<=1: return 'QUIT'
                                
                        choice_time = self.trialClock.getTime()-start_time
                if t>self.timer_limit: score,thisResp,thisResp_pos,choice_time = (0,'timed_out','timed_out','timed_out')

            #give feedback
            self.fb.present_fb(win,score,[self.stimulus,target_button,self.target]+foil_button+foil_text)

            #write data
            output = {
                'threshold_var': operation,
                'level': difficulty,
                'score': score,
                'resp_time': choice_time,
                'stim': stim_string,
                'resp': thisResp,
                'resp_pos': thisResp_pos,
                'target': target_string,
                'target_pos': pos[xpositions[0]],
                'foil1': foil1_string,
                'foil2': foil2_string,
                'foil3': foil3_string
            }
            
            i=1
            for name,foiltmp in zip(['foil1_pos','foil2_pos','foil3_pos'],[foil1_string,foil2_string,foil3_string]):
                if foiltmp!='':
                    output[name] = pos[xpositions[i]]
                    if total_foil==3: i+=1
                elif foiltmp=='': output[name] = ''
                
            if this_iteration[index] == len(these_conditions[index]['Correct'])-1: this_iteration[index] = 0
            else: this_iteration[index] += 1

            return output
