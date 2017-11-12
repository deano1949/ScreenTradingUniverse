%% Calculate all market beta of stocks in "New Universe" to SPY
% for VNM_V003_1 use

%% Load data
location='Home';

    if strcmp(location,'Home')
        addpath(genpath('C:\Users\Langyu\Desktop\Dropbox\GU\1.Investment\4. Alphas (new)'));
        path='C:\Users\Langyu\Desktop\Dropbox\GU\1.Investment\Data\NewUniverse\';
    elseif strcmp(location,'Coutts')
        addpath(genpath('O:\langyu\Reading\AlgorithmTrading_Chan_(2013)\jplv7'));
        addpath('O:\langyu\Reading\AlgorithmTrading_Chan_(2013)');
        path='O:\langyu\Reading\AlgorithmTrading_Chan_(2013)\NewUniverse\';
    else
        error('Unrecognised location; Coutts or Home');
    end

load(strcat(path, 'NUV.mat'));

idxStart=1;
idxEnd=length(cl);

cl=cl(idxStart:idxEnd, :); cl=fillMissingData(cl);
op=op(idxStart:idxEnd, :);cl=fillMissingData(cl);

%% Cal Market Beta
MIfile='C:\Users\Langyu\Desktop\Dropbox\GU\1.Investment\Data\NewUniverse\SPY_MarketIndex.csv';
spy=xlsread(MIfile);
spy=spy(idxStart:idxEnd,2); %load spy close price
spyret=(spy-backshift(1,spy))./backshift(1,spy); spyret(isnan(spyret))=0;
stckret=(cl-backshift(1,cl))./backshift(1,cl); stckret(1,:)=0;

beta=NaN(size(cl));
alpha=NaN(size(cl));
R11=NaN(size(cl));
R12=NaN(size(cl));
R22=NaN(size(cl));
for i=1:size(stckret,2)
    %initial beta (use first 252 pt to generate first beta
    fstpt=find(~isnan(stckret(:,i)) & stckret(:,i)~=0,1); %first nonzeros non-nan return value
    sndpt=fstpt+252-1;
    if sndpt<=size(cl,1)
        mdl=ols(stckret(fstpt:sndpt,i),[ones(252,1) spyret(fstpt:sndpt)]); fstbeta=mdl.beta;
        [B,~,~,R]=KalmanFilterG(spyret(sndpt+1:end),stckret(sndpt+1:end,i),0.001,fstbeta);
        beta(sndpt+1:end,i)=transpose(B(2,:));
        alpha(sndpt+1:end,i)=transpose(B(1,:));
        R11(sndpt+1:end,i)=transpose(R(1,1));
        R12(sndpt+1:end,i)=transpose(R(1,2));
        R22(sndpt+1:end,i)=transpose(R(2,2));
    end
end

Latestbeta=transpose(beta(end,:));
Latestalpha=transpose(alpha(end,:));
LatestR11=transpose(R11(end,:));
LatestR12=transpose(R12(end,:));
LatestR22=transpose(R22(end,:));

%% Output to xlsx
outputfile='outputfile.xlsx';
xlswrite(outputfile,{date},'Sheet2','A1');

xlswrite(outputfile,name','Sheet2','A5');
xlswrite(outputfile,{'alpha'},'Sheet2','B4');
xlswrite(outputfile,{'beta'},'Sheet2','C4');
xlswrite(outputfile,{'R11'},'Sheet2','D4');
xlswrite(outputfile,{'R12'},'Sheet2','E4');
xlswrite(outputfile,{'R22'},'Sheet2','F4');
xlswrite(outputfile,Latestalpha,'Sheet2','B5');
xlswrite(outputfile,Latestbeta,'Sheet2','C5');
xlswrite(outputfile,LatestR11,'Sheet2','D5');
xlswrite(outputfile,LatestR12,'Sheet2','E5');
xlswrite(outputfile,LatestR22,'Sheet2','F5');