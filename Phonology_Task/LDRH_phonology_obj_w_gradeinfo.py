from psychopy import gui, visual, core, data, event, logging, sound, info, misc
import time, numpy, os, sys, tempfile, wave
from os.path import join
from math import floor
from random import randint, choice, shuffle
from task_function import task_functions
if __name__ != '__main__': from Feedback import feedback
#touchscreen or not
touchscreen = True


class Phonology_Game(task_functions):

    def __init__(self, win, conditions):
        "Initialize the stimuli and import conditions"
        #get dir for importing resources
        self.dir = os.path.dirname(__file__)

        #directory holding stimuli, images and audio
        image_path = 'Images/Tasks/'
        audio_path = 'Audio/General/'
        aud_practice_path = 'Audio/Practice/'
        aud_inst_path = 'Audio/Instructions/'
        self.phonologystim_dir = 'Audio/Stimuli/Phonology/'

        #create practice instructions
        self.practice_cue1 = visual.TextStim(win, units=u'pix', wrapWidth=700, pos=[0,0],height=28,text="  Let's do some practice.\n\nTouch anywhere to begin.")
        self.practice_cue2 = visual.TextStim(win, units=u'pix', wrapWidth=700, pos=[0,0],height=28,text='Touch anywhere to do some more practice.')
        self.practice_cue3 = visual.TextStim(win, units=u'pix', wrapWidth=700, pos=[0,0],height=28,text="Are you ready to begin?")

        #initializing audio files for practice and instructions
        self.practice_aud1 = sound.Sound(aud_practice_path + 'practice_cue1.wav')
        self.practice_aud2 = sound.Sound(aud_practice_path + 'practice_cue2.wav')
        self.practice_aud3 = sound.Sound(aud_practice_path + 'practice_cue3.wav')
        self.phonology_inst1 = sound.Sound(aud_inst_path + 'phonology_inst1.wav')
        self.phonology_inst2 = sound.Sound(aud_inst_path + 'phonology_inst2.wav')
        self.phonology_inst3 = sound.Sound(aud_inst_path + 'phonology_inst3.wav')
        self.phonology_inst4 = sound.Sound(aud_inst_path + 'phonology_inst4.wav')

        #instructions
        self.instructions = visual.MovieStim(win=win,filename = aud_inst_path + 'phon_instructions.mp4', size = [1500,850], flipHoriz = True)
        self.audio_inst = sound.Sound(aud_inst_path + 'phon_instructions.wav')

        #create stimuli, repeat and continue button
        self.speaker = visual.ImageStim(win=win, name='speaker',image=image_path +'/speaker.png', mask = None, units=u'pix',ori=0, pos=[0,200], size=[115,115])
        self.speaker_playing = visual.ImageStim(win=win, name='speaker',units=u'pix',image=image_path +'/speaker_playing_white.png', mask = None,ori=0, pos=[45,200], size=[220,155])
        self.repeat_button=visual.ImageStim(win=win, name='repeat_button', image= image_path + 'repeat.png', units=u'pix', pos=[350, -300], size=[75,75], color=[1,1,1], colorSpace=u'rgb', opacity=1.0)
        self.continue_button=visual.ImageStim(win=win, name='continue_button', image= image_path + 'continue.png', units=u'pix', pos=[420, -300], size=[75,75], color=[1,1,1], colorSpace=u'rgb', opacity=1.0)
        self.same_button = visual.ImageStim(win, image=image_path + '/happy_button.png', pos=[-260, -200])
        self.different_button = visual.ImageStim(win, image=image_path + '/sad_button.png', pos=[260, -200])
        self.mouse=event.Mouse(win=win)
        self.mouse.getPos()
        self.trialClock = core.Clock()

        #time constrains
        self.t_initialspeaker = 1
        self.t_stimgap = 1
        self.timer_limit = 12

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


    def run_practice(self, win, grade):
        "Run practice"

        inst_set=[self.practice_cue1,None,None]
        aud_set=[self.practice_aud1,None,None]
        stim_set = [4,3,1]
        stim_repeat = stim_set
        var = ''
        score_cond = [None,None,None]
        
        return self.run_practice_functions(win, grade, inst_set, aud_set, stim_set, stim_repeat, score_cond, var)


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


    def run_game(self, win, grade, thisIncrement, var):
        "Run one iteration of the game with self.trialList as conditions."
        return self.run_trial(win, thisIncrement, var)

    def run_trial(self, win, thisIncrement, var):
        "Run one iteration of the game."
        self.trialClock.reset(); t=0

        #set the index to the current difficulty level for indexing into the conditions file
        for question in range(len(self.trialList)):
            print self.trialList[question].keys()
            if self.trialList[question]['Difficulty'] == (len(self.trialList)-thisIncrement):
                index = question

        def get_stims(stim):
            phonemes = [stim[x:x+2] for x in [0,2,4]]
            stim_files = [join(self.phonologystim_dir, phoneme.upper()+'.wav') for phoneme in phonemes]
            fn = join(self.temp_dir,'%s.wav'%stim)
            self.concat_wavs(stim_files, fn)
            audio = sound.Sound(value=fn)
            audio_length = audio.getDuration()
            return [audio,audio_length,stim]
            os.remove(fn)

        # Ensure iteration does not exceed length of available trials:
        if self.iteration[index] > len(self.trialList[index]['Stim1'])-1:
            self.iteration[index] = 0

        #check history to make sure we don't get more than three identical answers in a row; modify iteration if needed
        count=0 #give up after 50 tries
        while len(self.answer_history)>=3 and len(set(self.answer_history[-3:]))==1 and self.trialList[index]['Correct Response'][self.iteration[index]]==self.answer_history[-1] and count<50:
            if self.iteration[index] == len(self.trialList[index]['Stim1'])-1:
                self.iteration[index] = 0
            else:
                self.iteration[index] += 1
            count+=1

        #load trial variables
        difficulty = self.trialList[index]['Difficulty']
        stimA = self.trialList[index]['Stim1'][self.iteration[index]]
        stimB = self.trialList[index]['Stim2'][self.iteration[index]]
        target_content = self.trialList[index]['Correct Response'][self.iteration[index]]
        contents = ['same','different']
        contents.remove(target_content)
        foil_content = contents[0]
        print stimA, stimB, target_content

        #update answer_history
        self.answer_history.append(target_content)

        audioA = get_stims(stimA) #getstim() returns [audio,audio_length]
        audioB = get_stims(stimB)

        audio_order = [audioA,audioB]
        shuffle(audio_order)

        stim1 = audio_order[0][0]
        stim2 = audio_order[1][0]
        raw_stim1 = audio_order[0][2]
        raw_stim2 = audio_order[1][2]
        
        pos = {'same':['left',self.same_button], 'different':['right',self.different_button]}
        target_pos = pos[target_content][0]
        foil_pos = pos[foil_content][0]
        self.target_button = pos[target_content][1]
        self.foil_button = pos[foil_content][1]


        #display fixation with repeat, pause & continue button
        task_status = self.fixation_function()
        if task_status=='repeat_task': 
            print task_status
            return task_status

        elif task_status=='continue_task':
            print task_status

            #draw initial speaker before phoneme plays
            self.speaker.draw()
            win.flip()
            core.wait(self.t_initialspeaker)
            self.speaker_playing.draw()
            win.flip()
            
            #play first phoneme
            stim1.play()
            start_time = self.trialClock.getTime()
            while self.trialClock.getTime() < start_time + stim1.getDuration():
                if event.getKeys(keyList=['q', 'escape']): return 'QUIT'

            #after phoneme is played, wait one second and then play second tone
            self.speaker.draw()
            win.flip()
            core.wait(self.t_stimgap)
            self.speaker_playing.draw()
            win.flip()

            #play second phoneme
            stim2.play()
            start_time = self.trialClock.getTime()
            while self.trialClock.getTime() < start_time + stim2.getDuration():
                if event.getKeys(keyList=['q', 'escape']): return 'QUIT'

            self.speaker.draw()
            self.target_button.draw()
            self.foil_button.draw()
            win.flip()

            #start timer for response
            start_time=self.trialClock.getTime()
            choice_time=0
            thisResp=None
            thisResp_pos=None
            score = None
            self.mouse.getPos() #called to prevent last movement of mouse from triggering click
            while thisResp==None and choice_time<=self.timer_limit:
                if (self.mouse.mouseMoved() or (self.mouse.getPressed()==[1,0,0])):
                    if self.target_button.contains(self.mouse): score,thisResp,thisResp_pos = (1,target_content,target_pos)
                    elif self.foil_button.contains(self.mouse): score,thisResp,thisResp_pos = (0,foil_content,foil_pos)
                if event.getKeys(keyList=['escape']): return 'QUIT'
                choice_time=self.trialClock.getTime()-start_time

            if t>self.t_timer_limit: score,thisResp,thisResp_pos,choice_time = (0,'timed_out','timed_out','timed_out')    

            #give feedback
            self.fb.present_fb(win,score,[self.speaker,self.target_button,self.foil_button])

            #write data
            thresh = ['D3_PA-GA_KA-BA','D2_both_BA-TA_GA-TA_PA-DA_KA-DA','D1_POA__PA-TA_KA-TA_BA-DA_GA-DA','D2_POA_TA-DA_PA-KA_BA-GA_TA-DA','D1_VOT_PA-BA_KA-GA']
            threshold_var = thresh[difficulty-1]

            output = {
                'threshold_var': threshold_var,
                'level': difficulty,
                'score': score,
                'resp_time': choice_time,
                'stim1': raw_stim1,
                'stim2': raw_stim2,
                'resp': thisResp,
                'resp_pos': thisResp_pos,
                'target': target_content,
                'target_pos': target_pos,
            }
            
            output_header = ['phoneme_difference','POA_steps','VOT_steps','VOT_or_POA','phoneme_dif_pos','phoneme_dist']
            stim_header = ['Phoneme Difference','POA_steps','VOT_steps','VOT_or_POA','Difference Position','Distance']
            for out_col,stim_col in zip(output_header,stim_header):
                output.update({out_col:self.trialList[index][stim_col][self.iteration[index]]})
            
            #update iteration of current difficulty
            if self.iteration[index] == len(self.trialList[index]['Stim1'])-1:
                self.iteration[index] = 0
            else:
                self.iteration[index] += 1
            return output
            
