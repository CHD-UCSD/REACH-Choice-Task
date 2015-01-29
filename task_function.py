from psychopy import visual, core, data, event, logging, gui, sound
import os

class task_functions:
    
    def run_instructions(self, win):
        "Display the instructions for the game."
        #display instructions and wait
        self.audio_inst.play()
        while self.instructions._player.time <= int(self.instructions.duration):
            self.instructions.draw()
            win.flip()
        win.flip()

    def run_practice_functions(self, win, grade, inst_set, aud_set, stim_set, stim_repeat, score_cond, var):
        "Run practice"

        def run_sub_practice(self,win,text_cue,aud_cue,stim_condition,with_practice,option,score_cond,var):
            # self.repeat_button.draw() # self.continue_button.draw()
            if option=='no_repeat_option':
                if text_cue!=None and aud_cue!=None:
                    text_cue.draw()
                    aud_cue.play()
                    win.flip() #display instructions

                    #wait 1 seconds before checking for touch
                    start_time = self.trialClock.getTime()
                    while start_time+1 > self.trialClock.getTime():
                        if 'escape' in event.getKeys(): return 'QUIT'

                    #check for a touch
                    cont=False
                    self.mouse.getPos()
                    while cont==False:
                        if self.click(): 
                        	if aud_cue: aud_cue.stop()
                        	cont=True
                        if 'escape' in event.getKeys(): 
                        	if aud_cue: aud_cue.stop()
                        	return 'QUIT'
                else: win.flip()

            elif option=='repeat_opt':
                self.repeat_button.draw()
                self.continue_button.draw()
                text_cue.draw()
                aud_cue.play()
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
                            if aud_cue: aud_cue.stop()
                            return 'repeat'
                            break
                        elif self.continue_button.contains(self.mouse):
                            if aud_cue: aud_cue.stop()
                            return 'continue'
                            break
                    if 'escape' in event.getKeys():
                    	if aud_cue: aud_cue.stop()
                    	return 'QUIT'

            print 'with_practice', with_practice
            if with_practice==True:
                output = self.run_game(win, "", stim_condition,var)
                print 'run practice' #run first practice trial

                if score_cond:
                    while output['score']!=score_cond:
                        self.try_again.draw()
                        win.flip()
                        #wait a second before accepting touch
                        start_time=self.trialClock.getTime()
                        while start_time+1>self.trialClock.getTime():
                            if 'escape' in event.getKeys(): return 'QUIT'
                        #wait for a touch
                        self.mouse.getPos()
                        cont=False
                        while cont==False:
                            if self.click(): cont=True
                            if 'escape' in event.getKeys(): return 'QUIT'
                        output = self.run_game(win, "", stim_condition,var)

        def run_3_practice(inst,audio,stimuli,score_cond,var):
            #draw practice instructions, and do sub practice
            for txt,aud,stim in zip(inst,audio,stimuli,score_cond):
                run_sub_practice(self,win,txt,aud,stim,True,'no_repeat_option',score_cond,var)

        # for inst,aud,stim,score in zip (practice_var[0],practice_var[1],practice_var[2],practice_var[3]):
        if var = 'star_task':
            run_3_practice(inst_set[0],aud_set[0],stim_set[0],score_cond[0],var)
            run_3_practice(inst_set[1],aud_set[1],stim_set[1],score_cond[1],var)
            run_3_practice(inst_set[2],aud_set[2],stim_set[2],score_cond[2],var)
        else: run_3_practice(inst_set,aud_set,stim_set,score_cond,var)
        go_to_choice=False
        while go_to_choice==False:
            repeat_or_continue = run_sub_practice(self,win,self.practice_cue3,self.practice_aud3,None,False,'repeat_opt',var,score_cond)
            if repeat_or_continue=='repeat':
                if var=='star_task': run_3_practice(inst_set[3],aud_set[3],stim_repeat,score_cond[3],var)
                else: run_3_practice(inst_set,aud_set,stim_repeat,score_cond,var)
            elif repeat_or_continue=='continue':
                print 'continue2'
                go_to_choice=True
            if 'escape' in event.getKeys(): go_to_choice=True; return 'QUIT'

    