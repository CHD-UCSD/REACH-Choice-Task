from psychopy import gui, visual, core, data, event, logging, sound, info, misc
import time, numpy, os, sys, tempfile, wave
from os.path import join
from math import floor
from random import randint, choice
if __name__ != '__main__': from feedback import feedback
#touchscreen or not
touchscreen = True


class Phonology_Game:

    def __init__(self, win, conditions):
        "Initialize the stimuli and import conditions"
        #get dir for importing resources
        self.dir = os.path.dirname(__file__)

        #directory holding dtimuli
        self.stim_dir = 'final_phonemes'

        #create stimuli
        self.speaker = visual.ImageStim(win=win, name='speaker',image=self.dir +'/speaker.png', mask = None, units=u'pix',ori=0, pos=[0,200], size=[115,115])
        self.speaker_playing = visual.ImageStim(win=win, name='speaker',units=u'pix',image=self.dir +'/speaker_playing_white.png', mask = None,ori=0, pos=[45,200], size=[220,155])
        self.cue_touch = visual.TextStim(win, units=u'pix', pos=[0,-150],height=30, wrapWidth=700, text="Touch anywhere on the screen when you're ready to start.")
        self.instructions = visual.TextStim(win, units=u'pix', pos=[0,+100],height=30, wrapWidth=700, text="In this task, you will hear two made up words. Some of the time they will be the same, and some of the time they will be different. If the made up words are the same, touch the happy face button, and if they are not the same, touch the sad face button.")#  make your choice, touch the "Same" or "Different" button on the bottom of the screen.')
        self.trial_start = visual.TextStim(win, height=1, wrapWidth=25, pos=[0,+2], text='Ok, we\'re ready to start.')

        #create practice instructions and trial instructions
        self.practice_cue1 = visual.TextStim(win, units=u'pix', wrapWidth=700, pos=[0,0],height=28,text="         Let's do some practice.\n\n\n\nTouch anywhere to start practicing.")
        self.practice_cue2 = visual.TextStim(win, units=u'pix', wrapWidth=700, pos=[0,0],height=28,text='Touch anywhere to do some more practice.')
        self.practice_cue3 = visual.TextStim(win, units=u'pix', wrapWidth=700, pos=[0,0],height=28,text="Are you ready to begin?")
        self.practice_aud1 = sound.Sound('practice_cue1.wav')
        self.practice_aud2 = sound.Sound('practice_cue2.wav')
        self.practice_aud3 = sound.Sound('practice_cue3.wav')

        #repeat and continue button
        self.repeat_button=visual.ImageStim(win=win, name='repeat_button', image=u'repeat5.png', units=u'pix', pos=[350, -300], size=[75,75], color=[1,1,1], colorSpace=u'rgb', opacity=1.0)
        self.continue_button=visual.ImageStim(win=win, name='continue_button', image=u'continue5.png', units=u'pix', pos=[420, -300], size=[75,75], color=[1,1,1], colorSpace=u'rgb', opacity=1.0)

        #self.same_button = visual.ImageStim(win, image=self.fn + '/same_button_shapes.png', pos=[-260, -200])#, size=[300,140])
        #self.different_button = visual.ImageStim(win, image=self.fn + '/different_button_shapes.png', pos=[260, -200])#, size=[300,140])
        self.same_button = visual.ImageStim(win, image=self.dir + '/happy_button.png', pos=[-260, -200])
        self.different_button = visual.ImageStim(win, image=self.dir + '/sad_button.png', pos=[260, -200])
        self.mouse=event.Mouse(win=win)
        self.mouse.getPos()
        self.trialClock = core.Clock()

        #start feedback
        self.fb=feedback.fb(win)

        self.trialList=conditions

        #create a dictionary to keep track of how many times you've displayed each difficulty level
        self.iteration = {}
        for question in range(len(self.trialList)):
            self.iteration[question] = 0

        #list to keep track of history of answers
        self.answer_history = []

        #get tempdir
        self.temp_dir = tempfile.gettempdir()

    #method to get clicks
    def click(self):
        if touchscreen and self.mouse.mouseMoved(): return True
        elif not touchscreen and self.mouse.getPressed()==[1,0,0]: return True
        else: return False

    def run_instructions(self, win):
        "Display the instructions for the game."
        #display instructions and wait
        self.instructions.draw()
        self.cue_touch.draw()
        win.flip()
        #wait a second before checking for mouse movement
        core.wait(1)
        self.mouse.getPos()
        #check for a touch
        cont=False
        while cont==False:
            if self.click(): cont=True
            if 'escape' in event.getKeys(): core.quit()

    def run_practice(self, win, grade):
        "Run practice"

        def run_sub_practice(self,win,text_cue,aud_cue,stim_condition,with_practice,repeat_option):
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
            if with_practice==True: output = self.run_game(win, stim_condition); print 'run practice' #run first practice trial

        def run_3_practice(inst,audio,stimuli):
            #draw practice instructions, and do sub practice
            for txt,aud,stim in zip(inst,audio,stimuli):
                run_sub_practice(self,win,txt,aud,stim,True,'no_repeat_option')
            # run_sub_practice(self,win,self.practice_cue1,self.practice_aud1,3,True,'no_repeat_option')
            # run_sub_practice(self,win,self.practice_cue2,self.practice_aud2,2,True,'no_repeat_option')
            # run_sub_practice(self,win,self.practice_cue2,self.practice_aud2,1,True,'no_repeat_option')
        
        inst_set=[self.practice_cue1,self.practice_cue2,self.practice_cue2]
        aud_set=[self.practice_aud1,self.practice_aud2,self.practice_aud2]
        stim_set = [3,2,1]

        run_3_practice(inst_set,aud_set,stim_set)
        go_to_choice=False
        while go_to_choice==False:
            repeat_or_continue = run_sub_practice(self,win,self.practice_cue3,self.practice_aud3,None,False,'repeat_opt')
            if repeat_or_continue=='repeat': run_3_practice()
            elif repeat_or_continue=='continue':
                print 'continue2'
                go_to_choice=True
            if 'escape' in event.getKeys(): go_to_choice=True; return 'QUIT'


    def concat_wavs(self, infiles, outfile):
        data=[]
        for infile in infiles:
            w = wave.open(infile, 'rb')
            data.append( [w.getparams(), w.readframes(w.getnframes())] )
            w.close()
        output = wave.open(outfile, 'w')
        output.setparams(data[0][0])
        for i in range(len(infiles)):
            output.writeframes(data[i][1])
        output.close()


    def run_game(self, win, grade, thisIncrement):
        "Run one iteration of the game with self.trialList as conditions."
        return self.run_trial(win, thisIncrement)
    
    def run_trial(self, win, thisIncrement):
        "Run one iteration of the game."
        self.trialClock.reset(); t=0

        #set the index to the current difficulty level for indexing into the conditions file
        for question in range(len(self.trialList)):
            print self.trialList[question].keys()
            if self.trialList[question]['Difficulty'] == (len(self.trialList)-thisIncrement):
                index = question

        #load trial variables
        difficulty = self.trialList[index]['Difficulty']
        stim1 = self.trialList[index]['Stim1'][self.iteration[index]]
        stim2 = self.trialList[index]['Stim2'][self.iteration[index]]
        answer = self.trialList[index]['Correct Response'][self.iteration[index]]
        print stim1, stim2, answer

        #check history to make sure we don't get more than three identical answers in a row; modify iteration if needed
        count=0 #give up after 50 tries
        while len(self.answer_history)>=3 and len(set(self.answer_history[-3:]))==1 and answer==self.answer_history[-1] and count<50:
            if self.iteration[index] == len(self.trialList[index]['Stim1'])-1: self.iteration[index] = 0
            else: self.iteration[index] += 1
            count+=1

        #update answer_history
        self.answer_history.append(answer)

        #get stims
        phonemes1 = [stim1[x:x+2] for x in [0,2,4]]
        stim1_files = [join(self.dir, self.stim_dir, phoneme.upper()+'.wav') for phoneme in phonemes1]
        fn = join(self.temp_dir,'%s.wav'%stim1)
        self.concat_wavs(stim1_files, fn)
        audio1 = sound.Sound(value=fn)
        audio1_length = audio1.getDuration()
        os.remove(fn)

        phonemes2 = [stim2[x:x+2] for x in [0,2,4]]
        stim2_files = [join(self.dir, self.stim_dir, phoneme.upper()+'.wav') for phoneme in phonemes2]
        fn = join(self.temp_dir,'%s.wav'%stim2)
        self.concat_wavs(stim2_files, fn)
        audio2 = sound.Sound(value=fn)
        audio2_length = audio2.getDuration()
        os.remove(fn)

        #cue for start of trial
#        self.cue_touch.draw()
#        win.flip()
#        self.mouse.getPos()
#        while True:
#            if touchscreen and self.mouse.mouseMoved(): break
#            elif not touchscreen and self.mouse.getPressed()==[1,0,0]: break
#            if event.getKeys(keyList=['escape']): return 'QUIT'

        #draw the center dot
        self.speaker.draw()
        win.flip()

        #wait a second
        start_time = self.trialClock.getTime()
        while self.trialClock.getTime() < start_time + 1.0:
            if event.getKeys(keyList=['escape']): return 'QUIT'
        self.speaker_playing.draw()
        win.flip()

        #set order for trial
        order = choice([True,False])
        #check order and play first stimulus
        if order: 
            audio1.play()
            start_time=self.trialClock.getTime()
            while start_time+audio1_length > self.trialClock.getTime():
                if event.getKeys(keyList=['escape']): return 'QUIT'
        else: 
            audio2.play() 
            start_time=self.trialClock.getTime()
            while start_time+audio2_length > self.trialClock.getTime():
                if event.getKeys(keyList=['escape']): return 'QUIT'

        #after tone is played, wait one second and then play second tone
        self.speaker.draw()
        win.flip()
        start_time = self.trialClock.getTime()
        while self.trialClock.getTime() < start_time + 1.0:
            if event.getKeys(keyList=['escape']): return 'QUIT'
        self.speaker_playing.draw()
        win.flip()

        #play second stim
        if order: 
            audio2.play()
            start_time=self.trialClock.getTime()
            while start_time+audio2_length > self.trialClock.getTime():
                if event.getKeys(keyList=['escape']): return 'QUIT'
        else: 
            audio1.play() #play(stim1_files, audio1)
            start_time=self.trialClock.getTime()
            while start_time+audio1_length > self.trialClock.getTime():
                if event.getKeys(keyList=['escape']): return 'QUIT'

        #after the second tone has finished, put up the same and different buttons
        self.speaker.draw()
        self.same_button.draw()
        self.different_button.draw()
        win.flip()
        start_timer=self.trialClock.getTime()

        #wait for response
        thisResp=None
        self.mouse.getPos()
        while thisResp==None:
            click = self.click()
            #self.mouse.getPos()
            if click and self.same_button.contains(self.mouse): #screen click for "Same" button
                if answer == 'same': score, thisResp = (1,'same') #correct answer
                elif answer == 'different': score, thisResp = (0,'same') #incorrect answer
            elif click and self.different_button.contains(self.mouse): #screen click for "Different" button
                if answer == 'same': score, thisResp = (0, 'different') #incorrect answer
                elif answer == 'different': score, thisResp = (1, 'different') #correct answer
            if event.getKeys(keyList=['escape']): return 'QUIT'
        choice_time=self.trialClock.getTime()-start_timer

        #give feedback
        self.fb.present_fb(win,score,[self.speaker,self.same_button,self.different_button])

        #write data 
        #['Trial Number', 'Difficulty','Stim1','Stim2','Response','Correct Response','Score','Resp Time','POA_steps',
        # 'VOT_steps','VOT_or_POA','Difference Position','Distance','Number Phonemes','Phoneme Difference']
        #output = dict(self.trialList[index])
        output = {'Score':score,'Resp Time':choice_time,'Response':thisResp,'Stim1':stim1,'Stim2':stim2,'Correct Response':answer,'Difficulty':difficulty}
        for col in ['POA_steps','VOT_steps','VOT_or_POA','Difference Position','Distance','Number of Phonemes','Phoneme Difference']:
            output.update({col:self.trialList[index][col][self.iteration[index]]})

        #update iteration of current difficulty
        if self.iteration[index] == len(self.trialList[index]['Stim1'])-1: self.iteration[index] = 0
        else: self.iteration[index] += 1
        return output
