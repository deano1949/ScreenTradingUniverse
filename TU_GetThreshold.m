%% Get threshold for early exit for New Universe


%% Set up
ci_lvl=0.07; 
location='Home';

%%
if strcmp(location,'Home')
    addpath(genpath('C:\Users\Langyu\Desktop\Dropbox\GU\1.Investment\4. Alphas (new)'));
    addpath('C:\Users\Langyu\Desktop\Dropbox\GU\1.Investment\4. Alphas (new)\14.bog_sog')
    addpath('C:\Users\Langyu\Desktop\Dropbox\GU\1.Investment\4. Alphas (new)\14.bog_sog\1.Development\EarlyExit');

    path='C:\Spectrion\Data\NewUniverse\';
elseif strcmp(location,'Coutts')
    addpath(genpath('O:\langyu\Reading\AlgorithmTrading_Chan_(2013)\jplv7'));
    addpath('O:\langyu\Reading\AlgorithmTrading_Chan_(2013)');
    path='O:\langyu\Reading\AlgorithmTrading_Chan_(2013)\NewUniverse\';
else
    error('Unrecognised location; Coutts or Home');
end

load(strcat(path, 'NUV.mat'));

 Threshold = CH4_earlyexit(op,lo,hi,ci_lvl);
BogThreshold=transpose(Threshold.BOG(end,:));
SogThreshold=transpose(Threshold.SOG(end,:));

%% Output to xlsx
outputfile='outputfile.xlsx';
xlswrite(outputfile,{date},'Sheet1','A1');

xlswrite(outputfile,name','Sheet1','A5');
xlswrite(outputfile,{'BOG_Threshold'},'Sheet1','B4');
xlswrite(outputfile,{'SOG_Threshold'},'Sheet1','C4');
xlswrite(outputfile,BogThreshold,'Sheet1','B5');
xlswrite(outputfile,SogThreshold,'Sheet1','C5');

