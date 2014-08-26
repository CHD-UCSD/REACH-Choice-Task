# -*- coding: utf-8 -*-
from psychopy import core, visual, gui, data, misc, event, sound
import time, numpy, os, sys
from random import shuffle
if __name__ != '__main__': from Feedback import feedback

#touchscreen? if False, uses conventional mouse
touchscreen = True

#white rectangle instead of blue button
white_rectangle = False
#dark button instead of light blue button
dark_button = False

class Math_Game:
    
    def __init__(self, win, conditions):
        "Initialize the stimuli and iteration numbers, and import conditions"
        #get dir for importing conditions, images and audio
        self.dir = os.path.dirname(__file__)

        image_path = 'Images/Tasks/'
        audio_path = 'Audio/'
        math_dotstims_path = 'Images/Stimuli/Math_dotstims`'
        
        #create practice instructions
        self.practice_cue1 = visual.TextStim(win, units=u'pix', wrapWidth=700, pos=[0,0],height=28,text="         Let's do some practice.\n\n\n\nTouch anywhere to start practicing.")
        self.practice_cue2 = visual.TextStim(win, units=u'pix', wrapWidth=700, pos=[0,0],height=28,text='Touch anywhere to do some more practice.')
        self.practice_cue3 = visual.TextStim(win, units=u'pix', wrapWidth=700, pos=[0,0],height=28,text="Are you ready to begin?")
        self.practice_aud1 = sound.Sound('practice_cue1.wav')
        self.practice_aud2 = sound.Sound('practice_cue2.wav')
        self.practice_aud3 = sound.Sound('practice_cue3.wav')
        
        #repeat and continue button
        self.repeat_button=visual.ImageStim(win=win, name='repeat_button', image= image_path + 'repeat.png', units=u'pix', pos=[350, -300], size=[75,75], color=[1,1,1], colorSpace=u'rgb', opacity=1.0)
        self.continue_button=visual.ImageStim(win=win, name='continue_button', image= image_path + 'continue.png', units=u'pix', pos=[420, -300], size=[75,75], color=[1,1,1], colorSpace=u'rgb', opacity=1.0)
        
        #create stimuli
        self.foil1 = visual.TextStim(win, pos=[0,0],height=70, text='Foil1')
        self.foil2 = visual.TextStim(win, pos=[0,0],height=70, text='Foil2')
        self.foil3 = visual.TextStim(win, pos=[0,0],height=70, text='Foil3')
        self.foils = [self.foil1,self.foil2,self.foil3]
        self.target = visual.TextStim(win, pos=[0,0],height=70, text='Target.')
        self.text_stimulus = visual.TextStim(win, pos=[0,200],height=80, text='Stimulus.')
        self.dot_stimulus = visual.ImageStim(win,image=None,pos=[0,180],size=[260,260])
        self.fixation = visual.ImageStim(win, color='black', image=None, mask='circle',size=5)
        self.message1 = visual.TextStim(win, units=u'pix', pos=[0,+100], height=28, wrapWidth=700, text='In this game you will see a cluster of dots or a math problem at the top of the screen, and two or four possible answers at the bottom. Touch the answer you think is correct.')
        self.message2 = visual.TextStim(win, units=u'pix', pos=[0,-150],height=28, wrapWidth=700, text="Touch anywhere on the screen when you're ready to start.")
        self.foil1_button4 = visual.ImageStim(win, image= image_path + '/general_button_4.png')#, size=[300,120])
        self.foil2_button4 = visual.ImageStim(win, image= image_path + '/general_button_4.png')
        self.foil3_button4 = visual.ImageStim(win, image= image_path + '/general_button_4.png')
        self.target_button4 = visual.ImageStim(win, image= image_path + '/general_button_4.png')#, size=[300,120])
        self.foil1_button2 = visual.ImageStim(win, image= image_path + '/general_button.png')
        self.target_button2 = visual.ImageStim(win,image= image_path + '/general_button.png')

        self.mouse=event.Mouse(win=win)
        self.mouse.getPos()
        
        #start feedback
        self.fb=feedback.fb(win)
        
        self.trialList=conditions
        self.trialClock = core.Clock()
        
        #create a dictionary to keep track of how many times you've displayed each difficulty level
        self.iteration = {}
        for operation in ['addition','subtraction','multiplication','division']:
            self.iteration[operation] = {}
            for question in range(len(self.trialList[operation])):
                self.iteration[operation][question] = 0
        
    def run_instructions(self, win):
        "Display the instructions for the game."
        #display instructions and wait
        self.message1.draw()
        self.message2.draw()
        self.fixation.draw()
        win.flip()
        #wait a second before checking for mouse movement
        core.wait(1)
        self.mouse.getPos()
        #check for a touch
        cont=False
        while cont==False:
            if self.click(): cont=True
            if 'escape' in event.getKeys(): return 'QUIT'
    
    def run_practice(self, win):
        "Run practice"

        def run_sub_practice(self,win,text_cue,aud_cue,math_operation,stim_condition,with_practice,repeat_option):
            # self.repeat_button.draw() # self.continue_button.draw()
            if repeat_option=='no_repeat_option':
                text_cue.draw()
                # aud_cue.play()
                win.flip() #display instructions

                #wait 1 seconds before checking for touch
                start_time = self.trialClock.getTime()
                while start_time+1 > self.trialClock.getTime():
                    if 'escape' in event.getKeys(): return 'QUIT'
                
                #check for a touch
                cont=False
                self.mouse.getPos()
                while cont==False:
                    if self.click(): aud_cue.stop(); cont=True
                    if 'escape' in event.getKeys(): aud_cue.stop(); return 'QUIT'

            elif repeat_option=='repeat_opt': 
                self.repeat_button.draw()
                self.continue_button.draw()
                text_cue.draw()
                # aud_cue.play()
                win.flip() #display instructions

                #wait 1 seconds before checking for touch
                start_time = self.trialClock.getTime()
                while start_time+1 > self.trialClock.getTime():
                    if 'escape' in event.getKeys(): aud_cue.stop(); return 'QUIT'
                
                #check for a touch
                cont=False
                self.mouse.getPos()
                while cont==False:
                    if self.click():
                        if self.repeat_button.contains(self.mouse): #self.mouse.mouseMoved()
                            aud_cue.stop(); return 'repeat'
                            break
                        elif self.continue_button.contains(self.mouse):
                            aud_cue.stop(); return 'continue'
                            break
                    if 'escape' in event.getKeys(): aud_cue.stop(); return 'QUIT'
            
            print 'with_practice', with_practice
            if with_practice==True: output = self.run_game(win, math_operation, stim_condition); print 'run practice' #run first practice trial

        def run_3_practice(inst,audio,stimuli):
            #draw practice instructions, and do sub practice
            for txt,aud,stim in zip(inst,audio,stimuli):
                run_sub_practice(self,win,txt,aud,'addition',stim,True,'no_repeat_option')
            # run_sub_practice(self,win,self.practice_cue1,self.practice_aud1,'addition',13,True,'no_repeat_option')
            # run_sub_practice(self,win,self.practice_cue2,self.practice_aud2,'addition',11,True,'no_repeat_option')
            # run_sub_practice(self,win,self.practice_cue2,self.practice_aud2,'addition',11,True,'no_repeat_option')
        
        inst_set=[self.practice_cue1,self.practice_cue2,self.practice_cue2]
        aud_set=[self.practice_aud1,self.practice_aud2,self.practice_aud2]
        stim_set = [13,11,11]

        run_3_practice(inst_set,aud_set,stim_set)
        # run_3_practice()
        go_to_choice=False
        while go_to_choice==False:
            repeat_or_continue = run_sub_practice(self,win,self.practice_cue3,self.practice_aud3,None,None,False,'repeat_opt')
            if repeat_or_continue=='repeat': run_3_practice()
            elif repeat_or_continue=='continue':
                print 'continue2'
                go_to_choice=True
            if 'escape' in event.getKeys(): go_to_choice=True; return 'QUIT'

    def run_game(self, win, grade, thisIncrement):
        "Run one iteration of the game with self.trialList as conditions."
        return self.run_trial(win, operation, thisIncrement)
    
    def run_trial(self, win, operation, thisIncrement):
        "Run one iteration of the game."
        self.trialClock.reset()
        these_conditions = self.trialList[operation]
        this_iteration = self.iteration[operation]
        
        #set the index to the current difficulty level for indexing into the conditions file
        for question in range(len(these_conditions)):
            #'self.difficulty' increases in difficulty as numbers increase, thisIncrement increases in difficulty as numbers decrease
            if these_conditions[question]['Difficulty'] == (len(these_conditions)-thisIncrement):
                index = question
                difficulty = these_conditions[index]['Difficulty']
        print 'Difficulty is:', difficulty
        
        #randomize the position of the target and foil
        order= round(numpy.random.random())*2-1 #will be either +1(right) or -1(left)
        four_xpositions = [-360, -120, 120, 360]
        two_xpositions = [-260, 260]
        if these_conditions[index]['Foil2'] and these_conditions[index]['Foil3']: 
            number_options=4
            xpositions = four_xpositions
            target_button = self.target_button4
            foil_buttons = [self.foil1_button4, self.foil2_button4, self.foil3_button4]
            foils = [self.foil1,self.foil2,self.foil3]
        else: 
            number_options=2
            xpositions = two_xpositions
            target_button = self.target_button2
            foil_buttons = [self.foil1_button2]
            foils = [self.foil1]
        
        shuffle(xpositions)
        for xpos,button,object in zip(xpositions, [target_button]+foil_buttons, [self.target]+foils):
            object.setPos([xpos,-200])
            button.setPos([xpos,-200])
        
        #set the text for the stimuli
        stim_string = these_conditions[index]['Stimulus'][this_iteration[index]]

        #correct division signs if applicable
        if '*div' in stim_string: stim_string= stim_string.replace('*div',u'÷').replace('-',u'−')#stim_string[0:stim_string.index('*div')] + u'÷' + stim_string[stim_string.index('*div')+4:]
        #get image if applicable
        if '.png' in stim_string:
            self.dot_stimulus.setImage(math_dotstims_path+stim_string)
            self.stimulus = self.dot_stimulus
        else:
            self.text_stimulus.setText(stim_string)#[self.iteration[index]]))
            self.stimulus=self.text_stimulus
        
        # set target and foils
        target_string = str(these_conditions[index]['Correct'][this_iteration[index]])
        self.target.setText(target_string)#[self.iteration[index]]))
        self.target.setColor('white')
        
        foil1_string = str(these_conditions[index]['Foil1'][this_iteration[index]])
        self.foil1.setText(foil1_string)
        self.foil1.setColor('white')
        
        if number_options==4:
            foil2_string = str(these_conditions[index]['Foil2'][this_iteration[index]])
            self.foil2.setText(foil2_string)
            self.foil2.setColor('white')
            foil3_string = str(these_conditions[index]['Foil3'][this_iteration[index]])
            self.foil3.setText(foil3_string)
            self.foil3.setColor('white')
        else:
            foil2_string=''
            foil3_string=''
        
        for i,foil in enumerate(foils):
            foil.setText( str(these_conditions[index]['Foil{}'.format(i+1)][this_iteration[index]]) )#[self.iteration[index]]))
            foil.setColor('white')
        self.target.setText(str(these_conditions[index]['Correct'][this_iteration[index]]))#[self.iteration[index]]))
        self.target.setColor('white')
        
        #draw all stimuli
        self.stimulus.draw()
        target_button.draw()
        self.target.draw()
        self.fixation.draw()
        for button, foil in zip(foil_buttons, foils):
            button.draw()
            foil.draw()
        win.flip()
        
        #start the timer for the response
        start_time=self.trialClock.getTime()
        timer=0
        
        #get response
        score=None
        self.mouse.getPos() #called to prevent last movement of mouse from triggering click
        while score==None and timer<=12:
            if self.click():
                if target_button.contains(self.mouse):
                    score=1
                    self.target.setColor('gold')
                    response_chosen = self.target.text
                for button, foil in zip(foil_buttons, foils):
                    if button.contains(self.mouse):
                        score=0
                        foil.setColor('gold')
                        response_chosen = foil.text
            if event.getKeys(keyList=['q', 'escape']): return 'QUIT'
            timer=self.trialClock.getTime()-start_time
        #calculate response time
        if timer<12: response_time = timer
        else: 
            response_time = 'timed out'
            score=0
            response_chosen = None
        
        #give feedback
        self.fb.present_fb(win,score,[self.stimulus,target_button,self.target]+[foil_button for foil_button in foil_buttons]+[foil for foil in foils]) #[self.foil1_button,self.foil2_button,self.foil3_button,self.target_button,self.foil1,self.foil2, self.foil3,self.target])
        
        #write data #headers are ['Difficulty','Stimulus','Target','Foil','Score','Resp Time','Adaptive']
        output = {'Difficulty': difficulty, 'Stimulus': stim_string, 'Target': target_string, 'Foil1': foil1_string,'Foil2': foil2_string,'Foil3': foil3_string, 'Response Chosen': response_chosen, 'Score': score, 
            'Resp Time': response_time, 'Operation': operation}
        
        #update iteration of current difficulty
        if this_iteration[index] == len(these_conditions[index]['Correct'])-1: this_iteration[index] = 0
        else: this_iteration[index] += 1
        
        return output
        
    def end_game(self):
        "save files and exit game"
        #staircase has ended
#        dataFile.close()
        staircase.saveAsPickle(fileName)#special python binary file to save all the info
        staircase.saveAsExcel(fileName)
        
        
        #give some output to user
        #print 'reversals:'
        #staircase.reversalIntensities
        return 'reversals:', staircase.reversalIntensities, 'mean of final 6 reversals = %.3f' %(numpy.average(staircase.reversalIntensities[-6:]))
        
        #core.quit()
    
    #method to get clicks
    def click(self):
        if touchscreen and self.mouse.mouseMoved(): return True
        elif not touchscreen and self.mouse.getPressed()==[1,0,0]: return True
        else: return False


if __name__=='__main__':
    sys.path.append(os.path.abspath(os.path.join(os.getcwd(),os.pardir)))
    from Feedback import feedback
    
    win = visual.Window(size=(1100, 700), allowGUI=True, monitor=u'testMonitor', color=[-1,-1,-1], colorSpace=u'rgb', units=u'pix') #Window
    
    conditions = data.importConditions('generated_math_2and4_option_stims_no_zero.xlsx')
    staircase = data.StairHandler(startVal = 20, stepType = 'lin', stepSizes=[3,3,2,2,1,1], #reduce step size every two reversals
        minVal=0, maxVal=len(conditions)-1, nUp=1, nDown=3,  #will home in on the 80% threshold
        nTrials = 10)
    
    #initialize game
    game = Math_Game(win, conditions)
    
    #start feedback
    fb=feedback.fb(win)
    
    #step through staircase to find threshold
    for this_increment in staircase: 
        output = game.run_game(win, this_increment)
        staircase.addData(output['Score'])
    #record the resulting threshold level of the training
