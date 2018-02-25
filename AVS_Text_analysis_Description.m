filename = "Last 2 years\ICT - FS & AVS - Task&Customer - Last 2 years.xlsx";
T = readtable(filename);
head(T)
textData_AVS_Desc = T.description;
labels = T.assignment_group;

idx_AVS_D = labels == "ICT - Audio Visual Services" | labels == "ICT - AVS Educational Video" | labels == "ICT - AVS Engineering" | labels == "ICT - AVS Lecture Recording & Media Delivery" | labels == "ICT - AVS Video Conferencing & Web Conferencing";
textData_AVS_D = textData_AVS_Desc(idx_AVS_D);
labels_AVS_D = labels(idx_AVS_D);

cleanTextData_AVS_D = eraseTags(textData_AVS_D);
cleanTextData2_AVS_D = eraseURLs(cleanTextData_AVS_D);
cleanTextData3_AVS_D = erasePunctuation(cleanTextData2_AVS_D);
cleanTextData4_AVS_D = decodeHTMLEntities(cleanTextData3_AVS_D);
cleanTextData5_AVS_D = lower(cleanTextData4_AVS_D);

documents_AVS_D = tokenizedDocument(cleanTextData5_AVS_D);
cleanDocuments_AVS_D = removeWords(documents_AVS_D,stopWords);
cleanDocuments2_AVS_D = removeShortWords(cleanDocuments_AVS_D,2);

% Generate a Unigram Word Cloud
bag_AVS_D = bagOfWords(cleanDocuments2_AVS_D);
figure
wordcloud(bag_AVS_D)
title('AVS-Description Unclean Unigram word cloud')
topWords_U_AVS_D = topkwords(bag_AVS_D,200)

removal_words_AVS_D = ["2006";"2017";"2016";"2015";"please";"room";"contact";"unable";
    "level";"working";"like";"requires";"old";"today";"just";"9351";"minutes";
    "centre";"regards";"help";"request";"new";"set";"etc";"sent";"required";
    "number";"inroom";"subject";"sydney";"block";"problem";"get";"open";"time";
    "work";"currently";"last";"back";"monday";"tuesday";"wednesday";"thursday";
    "friday";"saturday";"sunday";"possible";"still";"message";"look";"come";
    "off";"seconds";"kind";"name";"week";"requested";"checked";"start";"needs";
    "dear";"nsw";"details";"able";"someone";"provide";"event";"same";"man";
    "email";"university";"support";"thanks";"tried";"many";"know";"general";"user";
    "ict";"link";"admin";"building";"experience";"via";"ensure";"information";
    "trying";"certain";"circumstances";"staff";"officer";"running";"thank"];

cleanBag_AVS_D = removeWords(bag_AVS_D,removal_words_AVS_D);
figure
wordcloud(cleanBag_AVS_D)
title('AVS-Description Clean Unigram word cloud')
topWords_C_AVS_D = topkwords(cleanBag_AVS_D,200)

% Generate a Bigram Word Cloud
f = @(s)s(1:end-1) + " " + s(2:end);
bigrams_AVS_D = docfun(f,cleanDocuments2_AVS_D);

bag2_AVS_D = bagOfWords(bigrams_AVS_D);
figure
wordcloud(bag2_AVS_D);
title('AVS-Description Unclean Bigram word cloud')
topWords2_U_AVS_D = topkwords(bag2_AVS_D,200)

removal_bigram_words_AVS_D = ["please contact";"minutes seconds";"trackit ticket";
    "support team";"conference set";"support available";"conference contact";
    "avs provide";"general avs";"provide inroom";"available conference";"set certain";
    "certain circumstances";"circumstances please";"team arrange";"arrange experience";
    "contact ext16000";"ext16000 option";"down minutes";"university sydney";
    "campus block";"check equipment";"please investigate";"create trackit";
    "ticket number";"theres issue";"start new";"form link";"cut past";"past ticket";
    "number google";"issue create";"ticket hashtag";"maintenance2017 open";
    "open tickets";"room requires";"ticket room";"room details";"enter room";
    "room checked";"room enter";"assign ticket";"devices room";"submit form";
    "requires switch";"link notes";"notes assigned";"assigned trackit";"checked cut";
    "form start";"ticket yourself";"yourself record";"record trackit";"trackit numbers";
    "numbers applicable";"applicable section";"section form";"form example";
    "example amx";"sending commands";"commands record";"record controltp";"controltp ticket";
    "ticket section";"section submit";"paul switch";"tickets unikey";"baro9905 assign";
    "msg got";"got questions";"questions suggestions";"details instructional";
    "hours minutes";"packet loss";"details check";"check boxes";"boxes inspect";
    "functionality theres";"room populate";"kind regards";"time msec";"please provide";
    "please assist";"let know";"far sites";"status down";"dns name";"please check";
    "status name";"device status";"down hours";"please fix";"pingecho time";
    "sydney nsw";"many thanks";"total attempts";"attempts shortterm";"shortterm packet";
    "last attempts";"attempts reset";"reset recent";"response time"; "tba tba";
    "time availability";"recent loss";"loss 100";"ict helpdesk";"nsw 2006";
    "investigate asap";"provide support";"support ict";"name unknown";"unknown address";
    "devices none";"none connected";"recent outages";"connected important";
    "outages reset";"msec recent"; "please attend";"please let";"requires urgent";
    "loss last";"please advise";"100 total";"helpdesk subject";"make test";
    "tba dns";"room complete";"room cut";"sent monday";"tickets own";"own unikey";
    "link start";"sent tuesday";"please note";"responding requires";"tape marked";
    "take look";"error message";"please confirm";"like someone";"attend asap";
    "room free";"sent wednesday";"hour minutes";"system room";"someone come";
    "attention requested";"ensure content";"dear ict";"plus attachments";"email plus";
    "receive email";"prohibited receive";"delete attachments";"attachments confidential";
    "strictly prohibited";"audio visual";"content sending";"sites support";
    "cricos 00026a";"host name";"next week";"okay interrupt";"availability hours";
    "support needed";"testing checking"; "interval minutes";"field calls";
    "please delete";"error please";"last updated";"happy interrupted";"presentation ensure";
    "doc sasha";"room completed";"mark room";"room mark";"last week";"please help";
    "ensure microphone";"required room";"schd test";"test required";"room please";
    "calling avs";"details theres";"sites hear";"seconds packet";"please arrange";
    "make sure";"inputs make";"check check";"check inputs";"complete echo";
    "echo schd";"please ensure";"sites need";"sites connect";"hear questions";
    "question time";"presenters aware";"appropriate echo";"100 last";"seconds last";
    "class starting";"please request";"audio video";"down hour";"ensure sites";
    "install appropriate";"default levels";"please perform";"someone please";
    "requesting assistance";"class room";"sent friday";"support required";
    "requires attention";"calls far";"technical support";"down unreachable";
    "return old";"please call";"turned off";"please visit";"campus assist";
    "sent thursday";"november 2017";"checking block";"every week";"please return";
    "minutes packet";"audio sources";"march 2017";"form theres";];

cleanBag2_AVS_D = removeWords(bag2_AVS_D,removal_bigram_words_AVS_D);
figure
wordcloud(cleanBag2_AVS_D);
title('AVS-Description Clean Bigram word cloud')
topWords2_C_AVS_D = topkwords(cleanBag2_AVS_D,200)

% Generate a Trigram Word Cloud
f = @(s)s(1:end-2) + " " + s(2:end-1) + " " + s(3:end);
trigrams_AVS_D = docfun(f,cleanDocuments2_AVS_D);

bag3_AVS_D = bagOfWords(trigrams_AVS_D);
figure
wordcloud(bag3_AVS_D);
title('FS-Description Unclean Trigram word cloud')
topWords3_U_AVS_D = topkwords(bag3_AVS_D,200)

removal_trigram_words_AVS_D = ["general avs provide";"avs provide inroom";
    "inroom support inroom";"support inroom support";"inroom support available";
    "support available conference";"available conference set";"conference set certain";
    "set certain circumstances";"certain circumstances please";"circumstances please contact";
    "support team arrange";"team arrange experience";"arrange experience connection";
    "conference contact ext16000";"contact ext16000 option";"faculty desktopcomputer support";
    "down minutes seconds";"performance check equipment";"create trackit ticket";
    "google form link";"cut past ticket";"past ticket number";"theres issue create";
    "issue create trackit";"maintenance2017 open tickets";"hashtag maintenance2017 open";
    "trackit ticket hashtag";"trackit ticket room";"complete maintenance google";
    "form link notes";"link notes assigned";"notes assigned trackit";"assigned trackit ticket";
    "ticket room checked";"room checked cut";"checked cut past";"room enter room";
    "assign ticket yourself";"ticket yourself record";"yourself record trackit";
    "record trackit numbers";"trackit numbers applicable";"numbers applicable section";
    "applicable section form";"section form example";"form example amx";"example amx sending";
    "amx sending commands";"sending commands record";"commands record controltp";
    "record controltp ticket";"controltp ticket section";"ticket section submit";
    "section submit form";"networked devices room";"devices room requires";
    "switch email paul";"open tickets unikey";"tickets unikey baro9905";"baro9905 assign ticket";
    "submit form msg";"form msg got";"msg got questions";"got questions suggestions";
    "switch details instructional";"details check boxes";"check boxes inspect";
    "boxes inspect rooms";"rooms functionality theres";"functionality theres issue";
    "form room populate";"device status name";"down hours minutes";"total attempts shortterm";
    "attempts shortterm packet";"shortterm packet loss";"last attempts reset";
    "attempts reset recent"; "response time msec";"reset recent loss";"university sydney nsw";
    "packet loss 100";"sydney nsw 2006";"please investigate asap";"support ict helpdesk";
    "name unknown address";"none connected important";"msec recent outages";
    "please provide support";"please let know";"time msec recent";"status down probe";
    "packet loss last";"loss last attempts";"down probe pingecho";"loss 100 total";
    "ict helpdesk subject";"tba dns name";"form room complete";"room complete echo";
    "echo check check";"check check inputs";"check inputs make";"inputs make test";
    "minutes seconds packet";"seconds packet loss";"room details theres";"details theres issue";
    "test recording populate";"echo schd test";"schd test required";"test required room";
    "video echo schd";"room mark room";"mark room completed";"room completed google";
    "doc sasha echo";"google doc sasha";"project presentation ensure";"hours minutes seconds";
    "interval minutes seconds";"availability hours minutes";"time availability hours";
    "street lidcombe nsw";"far sites hear";"msec minimum time"; "msec maximum time";
    "statistics response time";"echo statistics response";"heard well system";
    "audio sources heard";"audio confirm microphones";"microphones audio sources";
    "nsw 2006 9351";"minutes seconds last";"sites need technical";"need technical support";
    "sending correctly field";"content sending correctly";"correctly microphones unmuted";
    "ensure sites connect";"tms ensure sites";"monitor tms ensure";"questions monitor tms";
    "hear questions monitor";"sites hear questions";"time far sites";"question time far";
    "held mics question";"aware video conference";"ensure presenters aware";
    "needed ensure presenters";"support needed ensure";"many local regional";
    "local regional sites";"rounds videoconferenced many";"ccdevice monitor update";
    "appropriate echo software";"software ccdevice monitor";"down hour minutes";
    "seconds last updated";"sch grand rounds";"calls far sites";"status down unreachable";
    "field calls far";"field calls far";"lecture venue teaching";"testing checking block";
    "hours minutes packet"; "minutes packet loss";"form theres issue";"google form theres";
    "room cut past";"form room cut";"link start new";"open tickets own";"tickets own unikey";
    "please attend asap";"days hours minutes";"urgent attention requested";
    "confidential unauthorised strictly";"attachments confidential unauthorised";
    "00026a email plus";"cricos 00026a email";"ensure content sending";"strictly prohibited receive";
    "plus attachments confidential"; "email plus attachments";"error please delete";
    "receive email error";"tba tba dns";"email error please";"sites support needed"];

cleanBag3_AVS_D = removeWords(bag3_AVS_D,removal_trigram_words_AVS_D);
figure
wordcloud(cleanBag3_AVS_D);
title('AVS-Description Clean Trigram word cloud')
top3_C_AVS_D = topkwords(cleanBag3_AVS_D,200)

