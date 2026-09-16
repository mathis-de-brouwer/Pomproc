import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

const SEED = {"productCategories":[{"id":"c1","name":"EPI","order":1,"num":false,"color":"#E8792B","icon":"epi"},{"id":"c2","name":"Consommables","order":2,"num":false,"color":"#128A52","icon":"conso"},{"id":"c3","name":"Matériel","order":3,"num":false,"color":"#2D6CDF","icon":"mat"},{"id":"c4","name":"Outillage","order":4,"num":true,"color":"#0E9AA0","icon":"out"},{"id":"c5","name":"Machines","order":5,"num":true,"color":"#6C4BD6","icon":"mac"},{"id":"c6","name":"Carburant","order":6,"num":true,"color":"#D9453F","icon":"carb"},{"id":"c7","name":"Poubelles / sacs","order":7,"num":false,"color":"#64748B","icon":"pou"},{"id":"c8","name":"Nettoyage vitres","order":8,"num":false,"color":"#17A2C9","icon":"vit"},{"id":"c9","name":"Échelles / échafaudages","order":9,"num":true,"color":"#B8860B","icon":"ech"},{"id":"c10","name":"EPC","order":10,"num":false,"color":"#E0A400","icon":"epc"},{"id":"c11","name":"Textile","order":11,"num":false,"color":"#9B4DCA","icon":"tex","variants":[{"name":"Taille","options":["S","M","L","XL","XXL","XXXL"]},{"name":"Couleur","options":["Blanc","Noir","Bleu","Rouge"]}]}],"products":[{"id":"p1","catId":"c3","ref":"710522","nom":"Mop Uniko Microtex - 40 cm","unite":"Pièce","qteEmballage":25,"prix":7.71,"entity":"IIS","image":""},{"id":"p2","catId":"c3","ref":"710527","nom":"Mop Uniko Greenspeed Twist ABT - 40 cm","unite":"Pièce","qteEmballage":80,"prix":8.85,"entity":"IIS","image":""},{"id":"p3","catId":"c3","ref":"710520","nom":"Armature Uniko - 40 cm","unite":"Pièce","qteEmballage":10,"prix":27.06,"entity":"IIS","image":""},{"id":"p4","catId":"c3","ref":"710517","nom":"Mop triko Greenspeed Twist ABT - 40 cm","unite":"Pièce","qteEmballage":100,"prix":8.85,"entity":"IIS","image":""},{"id":"p5","catId":"c3","ref":"710518","nom":"Triko mop Scrub - 40 cm","unite":"Pièce","qteEmballage":25,"prix":8.59,"entity":"IIS","image":""},{"id":"p6","catId":"c3","ref":"710514","nom":"Mop Triko Microtex - 40 cm","unite":"Pièce","qteEmballage":25,"prix":7.96,"entity":"IIS","image":""},{"id":"p7","catId":"c3","ref":"710512","nom":"Armature Triko - 40 cm","unite":"Pièce","qteEmballage":10,"prix":22.58,"entity":"IIS","image":""},{"id":"p8","catId":"c3","ref":"706317","nom":"Mop velcro  Greenspeed Twist ABT - 30 cm","unite":"Pièce","qteEmballage":5,"prix":4.0,"entity":"IIS","image":""},{"id":"p9","catId":"c3","ref":"706017","nom":"Mop velcro Greenspeed Twist ABT - 45 cm","unite":"Pièce","qteEmballage":5,"prix":9.27,"entity":"IIS","image":""},{"id":"p10","catId":"c3","ref":"707111","nom":"Armature Plate pour Serpillière Velcro - 23 cm","unite":"Pièce","qteEmballage":10,"prix":16.0,"entity":"IIS","image":""},{"id":"p11","catId":"c3","ref":"707120","nom":"Armature velcro aluminium - 40 cm","unite":"Pièce","qteEmballage":10,"prix":17.6,"entity":"IIS","image":""},{"id":"p12","catId":"c2","ref":"771330","nom":"Greenspeed Mop2go mop jetable en microfibres - 45 x 13 cm - 10 pièces","unite":"Paquet","qteEmballage":20,"prix":5.05,"entity":"IIS","image":""},{"id":"p13","catId":"c3","ref":"111223","nom":"Frange de rechange balai à franges - 80 cm","unite":"Pièce","qteEmballage":25,"prix":11.5,"entity":"IIS","image":""},{"id":"p14","catId":"c3","ref":"732012C","nom":"ConceptCar 2010 Compact - avec roues standards","unite":"Pièce","qteEmballage":1,"prix":225.0,"entity":"IIS","image":""},{"id":"p15","catId":"c3","ref":"740668","nom":"Couvercle avec espace de rangement et support pour plan de travail","unite":"Pièce","qteEmballage":6,"prix":24.98,"entity":"IIS","image":""},{"id":"p16","catId":"c3","ref":"711330","nom":"MopCar 2 - 2 x 15 l - GRIS/VERT","unite":"Pièce","qteEmballage":1,"prix":85.0,"entity":"IIS","image":""},{"id":"p17","catId":"c3","ref":"706360","nom":"Presse à rouleau Concept - avec rouleau en position 'centrale' (anciennement: rouleau Ø 55 mm)","unite":"Pièce","qteEmballage":1,"prix":75.23,"entity":"IIS","image":""},{"id":"p18","catId":"c3","ref":"710522","nom":"Mop Uniko Microtex - 40 cm","unite":"Pièce","qteEmballage":25,"prix":7.71,"entity":"IIS","image":""},{"id":"p19","catId":"c2","ref":"350391","nom":"Lavette microfibres Micro Punto - 40 x 40 cm - BLEU - pack 5 pièces","unite":"Paquet","qteEmballage":48,"prix":3.0,"entity":"IIS","image":""},{"id":"p20","catId":"c2","ref":"350390","nom":"Lavette microfibres Micro Punto - 40 x 40 cm - ROSE - pack 5 pièces","unite":"Paquet","qteEmballage":48,"prix":3.0,"entity":"IIS","image":""},{"id":"p21","catId":"c2","ref":"350392","nom":"Lavette microfibres Micro Punto - 40 x 40 cm - JAUNE - paquet 5 pièces","unite":"Paquet","qteEmballage":48,"prix":3.0,"entity":"IIS","image":""},{"id":"p22","catId":"c2","ref":"350383","nom":"Lavette microfibres Micro Punto - 40 x 40 cm - VERT","unite":"Pièce","qteEmballage":10,"prix":0.6,"entity":"IIS","image":""},{"id":"p23","catId":"c2","ref":"131486","nom":"Gumclean Eponge Miracle - 11 x 6 x 3 cm","unite":"Pièce","qteEmballage":6,"prix":0.328,"entity":"IIS","image":""},{"id":"p24","catId":"c2","ref":"302351","nom":"Torchon microfibres Micro Floor - 53 x 70 cm - VERT","unite":"Pièce","qteEmballage":10,"prix":1.5,"entity":"IIS","image":""},{"id":"p25","catId":"c2","ref":"300356","nom":"Torchon non-tissé - 60 x 70 cm - ORANGE","unite":"Pièce","qteEmballage":100,"prix":0.7,"entity":"IIS","image":""},{"id":"p26","catId":"c2","ref":"130582","nom":"Éponge à récurer synthétique - 13 x 7 cm - BLEU/BLANC - 5 pièces","unite":"Paquet","qteEmballage":30,"prix":3.59,"entity":"IIS","image":""},{"id":"p27","catId":"c2","ref":"130581","nom":"Éponge à récurer synthétique - 13 x 7 cm - ROSE/BLANC - 5 pièces","unite":"Paquet","qteEmballage":30,"prix":3.59,"entity":"IIS","image":""},{"id":"p28","catId":"c2","ref":"130490","nom":"Éponge à récurer avec poignée Fabel - 10 x 7 cm - JAUNE/VERT","unite":"Pièce","qteEmballage":10,"prix":0.328,"entity":"IIS","image":""},{"id":"p29","catId":"c2","ref":"270422","nom":"Lingettes Oiltex - 60 x 22 cm - 1 x 100 pièces","unite":"Paquet","qteEmballage":12,"prix":3.65,"entity":"IIS","image":""},{"id":"p30","catId":"c3","ref":"140800","nom":"Manche métal plastifié avec filet - 150 cm","unite":"Pièce","qteEmballage":25,"prix":1.77,"entity":"IIS","image":""},{"id":"p31","catId":"c2","ref":"113154","nom":"Plumeau Magic Canari","unite":"Pièce","qteEmballage":10,"prix":2.4,"entity":"IIS","image":""},{"id":"p32","catId":"c3","ref":"176231B","nom":"Raclette à lame unique Boma Food - 40 cm - BLEU","unite":"Pièce","qteEmballage":10,"prix":18.08,"entity":"IIS","image":""},{"id":"p33","catId":"c3","ref":"750017","nom":"Vaporisateur sanitaire - 650 ml - ROUGE","unite":"Pièce","qteEmballage":1,"prix":2.15,"entity":"IIS","image":""},{"id":"p34","catId":"c3","ref":"750016","nom":"Vaporisateur intérieur - 650 ml","unite":"Pièce","qteEmballage":1,"prix":2.15,"entity":"IIS","image":""},{"id":"p35","catId":"c3","ref":"750019","nom":"Vaporisateur sol - 650 ml - VERT","unite":"Pièce","qteEmballage":1,"prix":2.15,"entity":"IIS","image":""},{"id":"p36","catId":"c2","ref":"270422","nom":"Lingettes Oiltex - 60 x 22 cm - 1 x 100 pièces","unite":"Paquet","qteEmballage":12,"prix":3.65,"entity":"IIS","image":""},{"id":"p37","catId":"c2","ref":"110422","nom":"Voiles imprégnés Oiltex - 60 x 22 cm - 12 x 100 pièces","unite":"Carton","qteEmballage":1,"prix":34.92,"entity":"IIS","image":""},{"id":"p38","catId":"c3","ref":"140800","nom":"Manche métal plastifié avec filet - 150 cm","unite":"Pièce","qteEmballage":25,"prix":1.77,"entity":"IIS","image":""},{"id":"p39","catId":"c3","ref":"120375","nom":"Manche en aluminium - 150 cm","unite":"Pièce","qteEmballage":10,"prix":4.15,"entity":"IIS","image":""},{"id":"p40","catId":"c3","ref":"107784","nom":"Raclette en plastique - 45 cm","unite":"Pièce","qteEmballage":10,"prix":7.85,"entity":"IIS","image":""},{"id":"p41","catId":"c3","ref":"120371","nom":"Manche en bois - 140 cm","unite":"Pièce","qteEmballage":25,"prix":1.9,"entity":"IIS","image":""},{"id":"p42","catId":"c2","ref":"710023","nom":"Pile alcaline type AA - 1,5 V - par  pièce","unite":"Pièce","qteEmballage":8,"prix":0.99,"entity":"IIS","image":""},{"id":"p43","catId":"c3","ref":"611384B","nom":"Pelle à poussière Boma - 26 cm - BLEU","unite":"Pièce","qteEmballage":60,"prix":2.98,"entity":"IIS","image":""},{"id":"p44","catId":"c3","ref":"097050","nom":"Balayette Boma Food douce - BLEU","unite":"Pièce","qteEmballage":10,"prix":6.0,"entity":"IIS","image":""},{"id":"p45","catId":"c3","ref":"113149","nom":"Tête de loup en crin avec manche télescopique 2 x 75 cm","unite":"Pièce","qteEmballage":12,"prix":6.6,"entity":"IIS","image":""},{"id":"p46","catId":"c3","ref":"113164","nom":"Tête-de-loup rond en PVC avec manche telescopique 84 - 150 cm","unite":"Pièce","qteEmballage":12,"prix":5.3,"entity":"IIS","image":""},{"id":"p47","catId":"c1","ref":"119315","nom":"Gants à usage unique en nitrile - non poudrés - BLEU - 100 pièces - SMALL","unite":"Boite","qteEmballage":10,"prix":4.98,"entity":"IIS","image":""},{"id":"p48","catId":"c1","ref":"119316","nom":"Gants à usage unique en nitrile - non poudrés - BLEU - 100 pièces - MEDIUM","unite":"Boite","qteEmballage":10,"prix":4.98,"entity":"IIS","image":""},{"id":"p49","catId":"c1","ref":"119317","nom":"Gants à usage unique en nitrile - non poudrés - BLEU - 100 pièces - LARGE","unite":"Boite","qteEmballage":10,"prix":4.98,"entity":"IIS","image":""},{"id":"p50","catId":"c1","ref":"119318","nom":"Gants à usage unique en nitrile - non poudrés - BLEU - 100 pièces - EXTRA LARGE","unite":"Boite","qteEmballage":10,"prix":4.98,"entity":"IIS","image":""},{"id":"p51","catId":"c1","ref":"119290","nom":"Gants de ménage en nitrile - BLEU - SMALL","unite":"Paire","qteEmballage":12,"prix":1.3,"entity":"IIS","image":""},{"id":"p52","catId":"c1","ref":"119291","nom":"Gants de ménage en nitrile - BLEU - MEDIUM","unite":"Paire","qteEmballage":12,"prix":1.3,"entity":"IIS","image":""},{"id":"p53","catId":"c1","ref":"119292","nom":"Gants de ménage en nitrile - BLEU - LARGE","unite":"Paire","qteEmballage":12,"prix":1.3,"entity":"IIS","image":""},{"id":"p54","catId":"c1","ref":"119293","nom":"Gants de ménage en nitrile - BLEU - EXTRA LARGE","unite":"Paire","qteEmballage":12,"prix":1.3,"entity":"IIS","image":""},{"id":"p55","catId":"c1","ref":"119390","nom":"Gants de ménage - BLEU - SMALL","unite":"Paire","qteEmballage":10,"prix":0.75,"entity":"IIS","image":""},{"id":"p56","catId":"c1","ref":"119391","nom":"Gants de ménage - BLEU - MEDIUM","unite":"Paire","qteEmballage":10,"prix":0.75,"entity":"IIS","image":""},{"id":"p57","catId":"c1","ref":"119392","nom":"Gants de ménage - BLEU - LARGE","unite":"Paire","qteEmballage":10,"prix":0.75,"entity":"IIS","image":""},{"id":"p58","catId":"c1","ref":"119325","nom":"Gants de ménage - ROUGE - MEDIUM","unite":"Paire","qteEmballage":10,"prix":0.75,"entity":"IIS","image":""},{"id":"p59","catId":"c1","ref":"119393","nom":"Gants de ménage - BLEU - EXTRA LARGE","unite":"Paire","qteEmballage":10,"prix":0.75,"entity":"IIS","image":""},{"id":"p60","catId":"c1","ref":"119427","nom":"Gants à usage unique en nitrile extra strong - non poudrés - NOIR - 100 pièces - LARGE","unite":"Boite","qteEmballage":10,"prix":8.9,"entity":"IIS","image":""},{"id":"p61","catId":"c1","ref":"119426","nom":"Gants à usage unique en nitrile extra strong - non poudrés - NOIR - 100 pièces - MEDIUM","unite":"Boite","qteEmballage":10,"prix":8.9,"entity":"IIS","image":""},{"id":"p62","catId":"c10","ref":"710009","nom":"Panneau de signalisation double 'Opgepast-Attention'","unite":"Pièce","qteEmballage":5,"prix":16.0,"entity":"IIS","image":""},{"id":"p63","catId":"c2","ref":"265021","nom":"Forcid détartrant puissant - 750 ml","unite":"Flacon","qteEmballage":16,"prix":3.03,"entity":"IIS","image":""},{"id":"p64","catId":"c2","ref":"291001","nom":"Fix Eco Swan - 750 ml","unite":"Flacon","qteEmballage":15,"prix":1.8,"entity":"IIS","image":""},{"id":"p65","catId":"c2","ref":"1326,1","nom":"Eau de javel précise - eau de javel épaissie  - 1 l","unite":"Flacon","qteEmballage":12,"prix":1.4,"entity":"IIS","image":""},{"id":"p66","catId":"c2","ref":"293021","nom":"Eco Swan WC détartrant - 750 ml","unite":"Flacon","qteEmballage":6,"prix":3.0,"entity":"IIS","image":""},{"id":"p67","catId":"c2","ref":"251041","nom":"Booster C - 1 l","unite":"Flacon","qteEmballage":10,"prix":4.1,"entity":"IIS","image":""},{"id":"p68","catId":"c2","ref":"250041","nom":"Booster C - 5 l","unite":"Bidon","qteEmballage":2,"prix":16.2,"entity":"IIS","image":""},{"id":"p69","catId":"c2","ref":"210245","nom":"Blitz Interior spray - 750 ml","unite":"Vapo","qteEmballage":12,"prix":2.52,"entity":"IIS","image":""},{"id":"p70","catId":"c2","ref":"210131","nom":"Fix Calc vinaigre de nettoyage - 1 l","unite":"Flacon","qteEmballage":12,"prix":1.17,"entity":"IIS","image":""},{"id":"p71","catId":"c2","ref":"210531","nom":"Vinaigre de nettoyage - 5 l","unite":"Bidon","qteEmballage":3,"prix":3.78,"entity":"IIS","image":""},{"id":"p72","catId":"c2","ref":"210171","nom":"DW Eco Rincit - 1 l (bientôt Dish Rinse)","unite":"Flacon","qteEmballage":12,"prix":4.84,"entity":"IIS","image":""},{"id":"p73","catId":"c2","ref":"220124","nom":"Autoscrub Pro 13 - 10 l","unite":"Bidon","qteEmballage":1,"prix":32.8,"entity":"IIS","image":""},{"id":"p74","catId":"c2","ref":"260150","nom":"Allygiene pamplemousse - 1 l","unite":"Flacon","qteEmballage":12,"prix":5.73,"entity":"IIS","image":""},{"id":"p75","catId":"c2","ref":"260550","nom":"Allygiene pamplemousse - 5 l","unite":"Bidon","qteEmballage":2,"prix":23.19,"entity":"IIS","image":""},{"id":"p76","catId":"c2","ref":"290004","nom":"Trio Floor 7 - 1 l","unite":"Flacon","qteEmballage":12,"prix":3.39,"entity":"IIS","image":""},{"id":"p77","catId":"c2","ref":"290504","nom":"Trio Floor 7 - 5 l","unite":"Bidon","qteEmballage":2,"prix":12.53,"entity":"IIS","image":""},{"id":"p78","catId":"c2","ref":"290008","nom":"Trio Interior 7 - 1 l","unite":"Flacon","qteEmballage":12,"prix":3.39,"entity":"IIS","image":""},{"id":"p79","catId":"c2","ref":"290508","nom":"Trio Interior 7 - 5 l","unite":"Bidon","qteEmballage":2,"prix":12.53,"entity":"IIS","image":""},{"id":"p80","catId":"c2","ref":"290002","nom":"Trio Sanitary 3 - 1 l","unite":"Flacon","qteEmballage":12,"prix":3.39,"entity":"IIS","image":""},{"id":"p81","catId":"c2","ref":"290502","nom":"Trio Sanitary 3 - 5 l","unite":"Bidon","qteEmballage":2,"prix":12.53,"entity":"IIS","image":""},{"id":"p82","catId":"c2","ref":"283176","nom":"Flacon doseur Eco Floor 6 non-remplissable 20 ml Dosy Mono - 1 l","unite":"Flacon","qteEmballage":10,"prix":4.8,"entity":"IIS","image":""},{"id":"p83","catId":"c2","ref":"293503","nom":"Eco Floor 6 - 5 l","unite":"Bidon","qteEmballage":2,"prix":14.8,"entity":"IIS","image":""},{"id":"p84","catId":"c2","ref":"284176","nom":"Flacon doseur Eco Floor 11 non-remplissable 20 ml Dosy Mono - 1 l","unite":"Flacon","qteEmballage":10,"prix":4.8,"entity":"IIS","image":""},{"id":"p85","catId":"c2","ref":"293304","nom":"Eco Floor 11 - 5 l","unite":"Bidon","qteEmballage":2,"prix":13.7,"entity":"IIS","image":""},{"id":"p86","catId":"c2","ref":"285176","nom":"Eco Sanitary 2 flacon doseur non-remplissable 20 ml Dosy Mono - 1 l","unite":"Flacon","qteEmballage":10,"prix":4.8,"entity":"IIS","image":""},{"id":"p87","catId":"c2","ref":"293502","nom":"Eco Sanitary 2 - 5 l","unite":"Bidon","qteEmballage":2,"prix":14.8,"entity":"IIS","image":""},{"id":"p88","catId":"c2","ref":"293501","nom":"Eco Interior 10 - 5 l","unite":"Bidon","qteEmballage":2,"prix":14.8,"entity":"IIS","image":""},{"id":"p89","catId":"c2","ref":"281176","nom":"Flacon doseur Eco Interior 10 non-remplissable 20 ml Dosy Mono - 1 l","unite":"Flacon","qteEmballage":10,"prix":5.38,"entity":"IIS","image":""},{"id":"p90","catId":"c2","ref":"214190","nom":"DW Eco Caps All In One capsules lave-vaisselle - 1,8 kg - 100 pièces (bientôt Dish Caps)","unite":"Boite","qteEmballage":5,"prix":17.93,"entity":"IIS","image":""},{"id":"p91","catId":"c2","ref":"298010","nom":"Fix Dish Vaisselle Citron - 1 l","unite":"Flacon","qteEmballage":15,"prix":1.37,"entity":"IIS","image":""},{"id":"p92","catId":"c2","ref":"210042","nom":"FD Soft détergent pour la vaisselle - 1 l (bientôt Dish Soft)","unite":"Flacon","qteEmballage":8,"prix":2.98,"entity":"IIS","image":""},{"id":"p93","catId":"c2","ref":"06134","nom":"Blackout - 750 ml","unite":"Vapo","qteEmballage":6,"prix":6.65,"entity":"IIS","image":""},{"id":"p94","catId":"c2","ref":"201803","nom":"Allcleaner - 1 l","unite":"Flacon","qteEmballage":10,"prix":3.6,"entity":"IIS","image":""},{"id":"p95","catId":"c2","ref":"260023","nom":"Déboucheur ménager - 1 l","unite":"Flacon","qteEmballage":12,"prix":3.6,"entity":"IIS","image":""},{"id":"p96","catId":"c2","ref":"201845","nom":"Iso Glass 10 nettoyant pour vitres - 1 l","unite":"Flacon","qteEmballage":10,"prix":3.37,"entity":"IIS","image":""},{"id":"p97","catId":"c2","ref":"333065","nom":"Tapinet No-tens R - 500 ml","unite":"Pièce","qteEmballage":12,"prix":10.17,"entity":"IIS","image":""},{"id":"p98","catId":"c2","ref":"210072","nom":"Sel adoucissant - 5 kg (bientôt Dish Salt)","unite":"Seau","qteEmballage":1,"prix":10.01,"entity":"IIS","image":""},{"id":"p99","catId":"c2","ref":"210172","nom":"Tablettes de sel - 25 kg","unite":"Sac","qteEmballage":1,"prix":20.71,"entity":"IIS","image":""},{"id":"p100","catId":"c2","ref":"210247","nom":"Blitz Interior - 5 l","unite":"Bidon","qteEmballage":2,"prix":9.0,"entity":"IIS","image":""},{"id":"p101","catId":"c2","ref":"210248","nom":"Blitz Forte spray - 750 ml","unite":"Vapo","qteEmballage":12,"prix":2.1,"entity":"IIS","image":""},{"id":"p102","catId":"c2","ref":"210249","nom":"Blitz Forte - 5 l","unite":"Bidon","qteEmballage":2,"prix":9.0,"entity":"IIS","image":""},{"id":"p103","catId":"c2","ref":"201813","nom":"Splendid - 1 l","unite":"Flacon","qteEmballage":10,"prix":3.2,"entity":"IIS","image":""},{"id":"p104","catId":"c2","ref":"205813","nom":"Splendid - 5 l","unite":"Bidon","qteEmballage":2,"prix":9.9,"entity":"IIS","image":""},{"id":"p105","catId":"c2","ref":"200046","nom":"XLG Probiotic Tab San - 300 ml - vide","unite":"Flacon","qteEmballage":110,"prix":6.8,"entity":"IIS","image":""},{"id":"p106","catId":"c2","ref":"200042","nom":"XLG PROBIOTIC Multi - 1 l","unite":"Flacon","qteEmballage":6,"prix":4.79,"entity":"IIS","image":""},{"id":"p107","catId":"c2","ref":"200040","nom":"XLG PROFORCE Swan - 750 ml","unite":"Flacon","qteEmballage":6,"prix":3.1,"entity":"IIS","image":""},{"id":"p108","catId":"c2","ref":"200044","nom":"XLG PROBIOTIC Duck - 750 ml","unite":"Flacon","qteEmballage":6,"prix":3.82,"entity":"IIS","image":""},{"id":"p109","catId":"c2","ref":"200043","nom":"XLG PROBIOTIC Multi - 5 l","unite":"Flacon","qteEmballage":2,"prix":19.28,"entity":"IIS","image":""},{"id":"p110","catId":"c2","ref":"200045","nom":"XLG Probiotic Tab Multi 300ml","unite":"Flacon","qteEmballage":110,"prix":6.8,"entity":"IIS","image":""},{"id":"p111","catId":"c2","ref":"200041","nom":"XLG PROFORCE Boost - 5 l","unite":"Bidon","qteEmballage":2,"prix":17.6,"entity":"IIS","image":""},{"id":"p112","catId":"c2","ref":"214200","nom":"Papier toilette traditionnel - tissu pur - 2 plis - 200 coupons - gaufré - BLANC - 48 rouleaux (12x4)","unite":"Paquet","qteEmballage":1,"prix":15.51,"entity":"IIS","image":""},{"id":"p113","catId":"c2","ref":"213200","nom":"Papier toilette traditionnel - tissu recyclé - 2 plis - 200 coupons - gaufré - BLANC - 64 rouleaux (16 x 4)","unite":"Paquet","qteEmballage":1,"prix":16.56,"entity":"IIS","image":""},{"id":"p114","catId":"c2","ref":"213400","nom":"Papier toilette traditionnel - tissu recyclé - 2 plis - 400 coupons - gaufré - BLANC - 40 rouleaux (10x4)","unite":"Paquet","qteEmballage":1,"prix":20.01,"entity":"IIS","image":""},{"id":"p115","catId":"c2","ref":"773311","nom":"Papier toilette Mini Jumbo - tissu recyclé - 2 plis - 180 m  - BLANC - 12 rouleaux","unite":"Paquet","qteEmballage":1,"prix":30.14,"entity":"IIS","image":""},{"id":"p116","catId":"c2","ref":"213350","nom":"Papier toilette Maxi Jumbo - tissu recyclé - 2 plis - 360 m  - BLANC - 6 rouleaux","unite":"Paquet","qteEmballage":1,"prix":22.69,"entity":"IIS","image":""},{"id":"p117","catId":"c2","ref":"774811","nom":"Rouleau essuie mains autocut - tissu pur - 2 plis - 150 m x 21 cm - BLEU - 6 rouleaux","unite":"Paquet","qteEmballage":1,"prix":66.0,"entity":"IIS","image":""},{"id":"p118","catId":"c2","ref":"774138","nom":"Essuie-mains Multifold - tissu pur - \"TAD\" - 2 plis - 32 x 20 cm - BLANC - 2000 pièces (20x100)","unite":"Carton","qteEmballage":1,"prix":43.77,"entity":"IIS","image":""},{"id":"p119","catId":"c2","ref":"774400","nom":"Essuie-mains Multifold - tissu pur - 2 plis - 32 x 20,3 cm - BLANC - 3000 pièces (25x120)","unite":"Paquet","qteEmballage":1,"prix":39.0,"entity":"IIS","image":""},{"id":"p120","catId":"c2","ref":"773200","nom":"Essuie-mains Singlefold - tissu recyclé - 2 plis - 23 x 24 cm - BLANC - 4000 pièces (20x200)","unite":"Paquet","qteEmballage":1,"prix":28.62,"entity":"IIS","image":""},{"id":"p121","catId":"c2","ref":"774200","nom":"Essuie-mains Singlefold - tissu pur - 2 plis - 24 x 21 cm - BLANC - 4000 pièces (20x200)","unite":"Carton","qteEmballage":1,"prix":30.49,"entity":"IIS","image":""},{"id":"p122","catId":"c2","ref":"773018","nom":"Essuie-mains Singlefold - tissu recyclé - 2 plis - 22 x 22.5 cm - NATUREL - 5000 pièces (20x250)","unite":"Carton","qteEmballage":1,"prix":26.7,"entity":"IIS","image":""},{"id":"p123","catId":"c2","ref":"773200A","nom":"Essuie-mains Singlefold - tissu recyclé - 2 plis - 22 x 22.5 cm - NATUREL - 5000 pièces (20x250)","unite":"Carton","qteEmballage":1,"prix":32.0,"entity":"IIS","image":""},{"id":"p124","catId":"c2","ref":"773105","nom":"Midi Centerfeed Multirol - tissu recyclé - 2 plis - 180 m x 20 cm - BLEU - 6 rouleaux","unite":"Paquet","qteEmballage":1,"prix":39.96,"entity":"IIS","image":""},{"id":"p125","catId":"c2","ref":"773004","nom":"Midi Centerfeed Multirol - tissu recyclé - 1 pli - 300 m x 19.5 cm - BLANC - 6 rouleaux","unite":"Paquet","qteEmballage":1,"prix":32.34,"entity":"IIS","image":""},{"id":"p126","catId":"c2","ref":"774001","nom":"Midi Centerfeed Multirol - tissu pur - 1 pli - 280 m x 20 cm - BLANC - 6 rouleaux","unite":"Paquet","qteEmballage":1,"prix":34.54,"entity":"IIS","image":""},{"id":"p127","catId":"c2","ref":"773313","nom":"Mini SmartOne papier toilette - 2 plis - 112 m x 13 cm - BLANC - 12 rouleaux - 472193","unite":"Paquet","qteEmballage":1,"prix":48.04,"entity":"IIS","image":""},{"id":"p128","catId":"c2","ref":"210008","nom":"Handsoap Mild - 5 l","unite":"Bidon","qteEmballage":2,"prix":8.6,"entity":"IIS","image":""},{"id":"p129","catId":"c2","ref":"214051","nom":"Essuie-tout - tissu pur - 2 plis - 50 coupons - gaufré - BLANC - 16 x 2 rouleaux","unite":"Paquet","qteEmballage":1,"prix":15.3,"entity":"IIS","image":""},{"id":"p130","catId":"c2","ref":"750162W","nom":"Tork Premium Mild Foam Soap - 1 l - 6 pièces","unite":"Carton","qteEmballage":1,"prix":110.88,"entity":"IIS","image":""},{"id":"p131","catId":"c2","ref":"770153","nom":"Tork Premium Soap Liquid Extra Mild Non Perfumed - 1L - 420701","unite":"Flacon","qteEmballage":6,"prix":10.4,"entity":"IIS","image":""},{"id":"p132","catId":"c2","ref":"774830","nom":"Admire Foam Soap - 1 l","unite":"Pièce","qteEmballage":6,"prix":6.87,"entity":"IIS","image":""},{"id":"p133","catId":"c2","ref":"261034","nom":"Tapis urinoir - Pacco - GRIS","unite":"Pièce","qteEmballage":10,"prix":2.6,"entity":"IIS","image":""},{"id":"p134","catId":"c3","ref":"114157","nom":"Brosse de toilette en support avec frotteur rebord - BLANC","unite":"Pièce","qteEmballage":32,"prix":4.04,"entity":"IIS","image":""},{"id":"p135","catId":"c2","ref":"774002","nom":"Mini Centerfeed Multirol - tissu pur - 1 pli - 120 m x 20 cm - BLANC - 12 rouleaux","unite":"Paquet","qteEmballage":1,"prix":34.3,"entity":"IIS","image":""},{"id":"p136","catId":"c7","ref":"625343","nom":"Sac HD \"universel\" - 70 x 110 cm - T20 - 115 l - Blanc - rouleau 25 pièces","unite":"Rouleau","qteEmballage":10,"prix":2.87,"entity":"IIS","image":""},{"id":"p137","catId":"c7","ref":"625342","nom":"Sac HD \"universel\" - 70 x 110 cm - T20 - 115 l - TRANSPARENT - rouleau 25 pièces","unite":"Rouleau","qteEmballage":10,"prix":3.84,"entity":"IIS","image":""},{"id":"p138","catId":"c7","ref":"625362","nom":"Sac HD \"universel\" - 70 x 110 cm - T20 - 115 l - GRIS - rouleau 25 pièces","unite":"Rouleau","qteEmballage":10,"prix":2.03,"entity":"IIS","image":""},{"id":"p139","catId":"c7","ref":"625377","nom":"Sac HD \"universel\" - 70 x 110 cm - T20 - 115 l - BLEU - rouleau 25 pièces","unite":"Rouleau","qteEmballage":10,"prix":3.66,"entity":"IIS","image":""},{"id":"p140","catId":"c7","ref":"625385","nom":"Sac BD - 70 x 110 cm - T45 - 115 l - JAUNE - rouleau 25 pièces","unite":"Rouleau","qteEmballage":10,"prix":4.6,"entity":"IIS","image":""},{"id":"p141","catId":"c7","ref":"625383","nom":"Sac BD - 70 x 110 cm - T45 - 115 l - BLEU - rouleau 25 pièces","unite":"Rouleau","qteEmballage":10,"prix":4.82,"entity":"IIS","image":""},{"id":"p142","catId":"c7","ref":"625384","nom":"Sac BD - 70 x 110 cm - T45 - 115 l - NOIR - rouleau 25 pièces","unite":"Rouleau","qteEmballage":10,"prix":4.82,"entity":"IIS","image":""},{"id":"p143","catId":"c7","ref":"650348","nom":"Sac HD - 50 x 55 cm - T10 - 25 l - NOIR - rouleau 50 pièces","unite":"Rouleau","qteEmballage":20,"prix":1.02,"entity":"IIS","image":""},{"id":"p144","catId":"c7","ref":"650367","nom":"Sac HD - 60 x 60 cm - T10 - 35 l - NOIR - rouleau 50 pièces","unite":"Rouleau","qteEmballage":20,"prix":0.91,"entity":"IIS","image":""},{"id":"p145","catId":"c7","ref":"650372","nom":"Sac HD - 60 x 60 cm - T10 - 35 l - TRANSPARENT - rouleau 50 pièces","unite":"Rouleau","qteEmballage":20,"prix":1.05,"entity":"IIS","image":""},{"id":"p146","catId":"c7","ref":"625703","nom":"Sac HD - 90 x 120 cm - T40 - 195 l - GRIS - rouleau 25 pièces","unite":"Rouleau","qteEmballage":10,"prix":6.94,"entity":"IIS","image":""},{"id":"p147","catId":"c7","ref":"625354","nom":"Sac HD - 120 x 140 cm - T40 - 365 l - GRIS - rouleau 15 pièces","unite":"Rouleau","qteEmballage":8,"prix":8.66,"entity":"IIS","image":""},{"id":"p148","catId":"c8","ref":"505227","nom":"Housse mouilleur en microfibres Boma - 25 cm - BLUE","unite":"Pièce","qteEmballage":10,"prix":4.8,"entity":"IIS","image":""},{"id":"p149","catId":"c8","ref":"505228","nom":"Housse mouilleur en microfibres Boma - 35 cm - BLEU","unite":"Pièce","qteEmballage":10,"prix":5.78,"entity":"IIS","image":""},{"id":"p150","catId":"c8","ref":"505229","nom":"Housse mouilleur en microfibres Boma - 45 cm - BLEU","unite":"Pièce","qteEmballage":10,"prix":6.98,"entity":"IIS","image":""},{"id":"p151","catId":"c8","ref":"500252","nom":"Support T Boma - 25 cm","unite":"Pièce","qteEmballage":100,"prix":2.98,"entity":"IIS","image":""},{"id":"p152","catId":"c8","ref":"500253","nom":"Support T Boma - 35 cm","unite":"Pièce","qteEmballage":100,"prix":3.28,"entity":"IIS","image":""},{"id":"p153","catId":"c8","ref":"500254","nom":"Support T Boma - 45 cm","unite":"Pièce","qteEmballage":50,"prix":3.68,"entity":"IIS","image":""},{"id":"p154","catId":"c8","ref":"500301","nom":"Raclette vitres ménagère - 25 cm","unite":"Pièce","qteEmballage":25,"prix":1.67,"entity":"IIS","image":""},{"id":"p155","catId":"c8","ref":"503210","nom":"Raclette vitres en inox Boma - 25 cm","unite":"Pièce","qteEmballage":1,"prix":10.6,"entity":"IIS","image":""},{"id":"p156","catId":"c8","ref":"503212","nom":"Raclette vitres en inox Boma - 35 cm","unite":"Pièce","qteEmballage":1,"prix":11.3,"entity":"IIS","image":""},{"id":"p157","catId":"c8","ref":"503213","nom":"Raclette vitres en inox Boma - 45 cm","unite":"Pièce","qteEmballage":1,"prix":12.0,"entity":"IIS","image":""},{"id":"p158","catId":"c8","ref":"500216","nom":"Caoutchouc de rechange Boma - 25 cm","unite":"Pièce","qteEmballage":50,"prix":1.29,"entity":"IIS","image":""},{"id":"p159","catId":"c8","ref":"500218","nom":"Caoutchouc de rechange Boma - 35 cm","unite":"Pièce","qteEmballage":50,"prix":1.69,"entity":"IIS","image":""},{"id":"p160","catId":"c8","ref":"500219","nom":"Caoutchouc de rechange Boma - 45 cm","unite":"Pièce","qteEmballage":50,"prix":1.9,"entity":"IIS","image":""},{"id":"p161","catId":"c4","ref":"701306","nom":"Aspirateur BOMA Aspiro Dry 11 Plus","unite":"Pièce","qteEmballage":1,"prix":165.0,"entity":"IIS","image":""},{"id":"p162","catId":"c4","ref":"Z701305","nom":"Aspirateur BOMA Aspiro Dry 11 Recycled","unite":"Pièce","qteEmballage":1,"prix":110.0,"entity":"IIS","image":""},{"id":"p163","catId":"c2","ref":"106013","nom":"Set de sacs à poussière Fleece Aspiro Dry 11 (Plus) 10pcs","unite":"Set","qteEmballage":20,"prix":8.4,"entity":"IIS","image":""},{"id":"p164","catId":"c4","ref":"OUT-001","nom":"Monobrosse 43 cm","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p165","catId":"c4","ref":"OUT-002","nom":"Petit Karcher / Kranzl HP compact","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p166","catId":"c4","ref":"OUT-003","nom":"Pompe immergée eaux usées","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p167","catId":"c4","ref":"OUT-004","nom":"Aspirateur eau et poussière","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p168","catId":"c5","ref":"MAC-001","nom":"Auto-laveuse autoportée","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p169","catId":"c5","ref":"MAC-002","nom":"Karcher 500 bar","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":"","assoc":["p186"]},{"id":"p170","catId":"c5","ref":"MAC-003","nom":"Balayeuse autoportée","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p171","catId":"c5","ref":"MAC-004","nom":"Groupe électrogène","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":"","assoc":["p173","p175"]},{"id":"p172","catId":"c6","ref":"CAR-001","nom":"Bidon 5 L mélange 2 temps","unite":"Bidon","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p173","catId":"c6","ref":"CAR-002","nom":"Bidon 20 L essence","unite":"Bidon","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p174","catId":"c6","ref":"CAR-003","nom":"Bidon 10 L diesel","unite":"Bidon","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p175","catId":"c6","ref":"CAR-004","nom":"Bidon 20 L huile","unite":"Bidon","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p176","catId":"c9","ref":"ECH-001","nom":"Échelle télescopique 3 m","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p177","catId":"c9","ref":"ECH-002","nom":"Échelle double 2×8","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p178","catId":"c9","ref":"ECH-003","nom":"Échafaudage roulant","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p179","catId":"c10","ref":"EPC-001","nom":"Cône de signalisation","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p180","catId":"c10","ref":"EPC-002","nom":"Rubalise 100 m","unite":"Rouleau","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p181","catId":"c10","ref":"EPC-003","nom":"Barrière de chantier","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p182","catId":"c11","ref":"TEX-001","nom":"T-shirt XLG","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p183","catId":"c11","ref":"TEX-002","nom":"Polo XLG","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p184","catId":"c11","ref":"TEX-003","nom":"Pantalon de travail","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p185","catId":"c11","ref":"TEX-004","nom":"Veste de travail","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""},{"id":"p186","catId":"c3","ref":"TUY-001","nom":"Tuyau d'eau 20 m","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":"","assoc":["p187"]},{"id":"p187","catId":"c3","ref":"RAC-001","nom":"Raccord Gardena","unite":"Pièce","qteEmballage":1,"prix":0.0,"entity":"IIS","image":""}],"chantierCategories":[],"templates":[{"id":"tpl_bureaux","name":"Bureaux","ownerUsername":"admin","shared":true,"lines":[{"productId":"p63","qty":2,"mode":"unite"},{"productId":"p18","qty":4,"mode":"unite"},{"productId":"p19","qty":3,"mode":"unite"},{"productId":"p112","qty":1,"mode":"unite"},{"productId":"p136","qty":2,"mode":"unite"},{"productId":"p47","qty":1,"mode":"unite"}],"createdAt":0},{"id":"tpl_evenementiel","name":"Événementiel","ownerUsername":"admin","shared":true,"lines":[{"productId":"p47","qty":2,"mode":"unite"},{"productId":"p136","qty":4,"mode":"unite"},{"productId":"p63","qty":2,"mode":"unite"},{"productId":"p179","qty":6,"mode":"unite"}],"createdAt":0},{"id":"tpl_vitres","name":"Vitres","ownerUsername":"admin","shared":true,"lines":[{"productId":"p148","qty":3,"mode":"unite"},{"productId":"p152","qty":2,"mode":"unite"},{"productId":"p154","qty":2,"mode":"unite"},{"productId":"p96","qty":3,"mode":"unite"}],"createdAt":0}],"counters":{}};
const LOGO = "data:image/png;base64,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WSB+Y6f7pPjteRwHqB+vzKQdUXeKx2Se4S4LmDEbD8d55D/+7AVAFVPLU1MtTO8vlleXvceZJOSVm6Na7c3WlwXDtK37Rtc+j2y0+k/invl1R6O9+SfSfNERWY0mEREAWSr7xVVlnobU9kUdPR7xYGAguLjkl3Hif4lY1F1lBSabXA9adepSjKMHhSWH1rKePFIIiLseQREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBSjsl00YYxf62PEkjcUrSOTTzf6+Q82e9axs60z7v3MzVI/kNMQZfxz2M/j5vSpeobhBUV9RQ0kRMdIAySQYDGv+QO8gc+7goPVrxqLo0+/qXQbP5A8noSqw1G64ZxTXS1xfYsPHXv5j3oi03afqT3ItfgFLJiuqmkAg8Y2ci70nkPX3Kv0KMq1RQjxZtrU9Ro6bazuqz+GK8XzJdbNK2oah917x4HTPzR0ZLWkHg9/xnf3D19609EV3oUY0aahHgj5l1PUK2pXU7qs/ik/DoXYluCIi9TACIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiA91pulxtrntoK+akE2GyFjiAfOfR3qbNMss9otlPb6a508z3ne3zM0ume7m4ce1QKg4HIWBeWKuVjOO7j/wDC1cnOVE9Fk5On7zdhZbWys5eOK37s9hYTUF6o7NaJbjUPDmN8VjWni9/Y0ev2YKge8XGputymr6t+9LK7J7mjsA8wC87ppnQthdK8xNJc1hcd0E8yAvmuLHT42uXnLfP1HflPysra64R2diEf5c539L4di/3CIikCpBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAf/Z";

const KEY_DATA = "pomproc:data3";
const KEY_ORDERS = "pomproc:orders3";
async function sget(k) { try { const r = await window.storage.get(k, true); return r ? JSON.parse(r.value) : null; } catch (e) { return null; } }
async function sset(k, v) { try { await window.storage.set(k, JSON.stringify(v), true); return true; } catch (e) { console.error(e); return false; } }

const ENTITIES = { FFA: { label: "FFA", color: "#0E9AA0", bill: true }, IIS: { label: "IIS", color: "#128A52", bill: false }, DC: { label: "DC", color: "#D9453F", bill: true } };
const ENT_ORDER = ["FFA", "IIS", "DC"];
const STATUS = { envoyee: { label: "À préparer", color: "#C9992B" }, en_cours: { label: "En préparation", color: "#2D6CDF" }, terminee: { label: "Terminée", color: "#128A52" } };
const ROLES = [{ id: "user", label: "Demandeur" }, { id: "magasinier", label: "Magasinier" }, { id: "admin", label: "Administrateur" }];
const APP_VERSION = "v3.1";

const money = (n) => (Number(n) || 0).toFixed(2).replace(".", ",") + " €";
const uid = (p) => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const byName = (a, b) => a.nom.localeCompare(b.nom, "fr", { sensitivity: "base" });
const monthKey = (ts) => { const d = new Date(ts); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0"); };
const monthLabel = (m) => { const [y, mo] = m.split("-"); return new Date(y, mo - 1, 1).toLocaleDateString("fr-BE", { month: "long", year: "numeric" }); };
const fmtD = (s) => { if (!s) return "—"; const d = new Date(s.length <= 10 ? s + "T00:00" : s); return isNaN(d) ? "—" : d.toLocaleDateString("fr-BE"); };
function lineTotal(l, p) { if (!p) return 0; const f = l.mode === "paquet" ? (p.qteEmballage || 1) : 1; return l.qty * f * p.prix; }
function lineKey(l) { return l.variant ? l.productId + "::" + l.variant : l.productId; }
function normGroups(v) { if (!v || !v.length) return []; if (typeof v[0] === "string") return [{ name: "Variante", options: v }]; return v.map((g) => ({ name: g.name || "Variante", options: g.options || [] })).filter((g) => g.options.length); }
function assocClosure(startId, resolve) { const seen = new Set([startId]); const out = []; const queue = [...(resolve(startId)?.assoc || [])]; while (queue.length) { const id = queue.shift(); if (seen.has(id)) continue; seen.add(id); const prod = resolve(id); if (!prod) continue; out.push(id); (prod.assoc || []).forEach((a) => { if (!seen.has(a)) queue.push(a); }); } return out; }
function variantGroupsOf(p, catById) { const pv = normGroups(p?.variants); if (pv.length) return pv; return normGroups(catById[p?.catId]?.variants); }
function parseGroups(text) { return String(text || "").split(/\n+/).map((line) => { line = line.trim(); if (!line) return null; let name = "Variante", rest = line; const ci = line.indexOf(":"); if (ci >= 0) { name = line.slice(0, ci).trim() || "Variante"; rest = line.slice(ci + 1); } const options = rest.split(",").map((s) => s.trim()).filter(Boolean); return options.length ? { name, options } : null; }).filter(Boolean); }
function groupsToText(v) { return normGroups(v).map((x) => x.name + ": " + x.options.join(", ")).join("\n"); }
function downloadBlob(blob, name) { const u = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = u; a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(u), 1500); }

/* ---- Notifications : ouvre WhatsApp / e-mail pré-remplis (envoi manuel) ---- */
const onlyDigits = (s) => String(s || "").replace(/[^\d]/g, "");
function waLink(phone, text) { const d = onlyDigits(phone); return d ? "https://wa.me/" + d + "?text=" + encodeURIComponent(text) : ""; }
function mailLink(emails, subject, body) { const to = Array.isArray(emails) ? emails.filter(Boolean).join(",") : (emails || ""); return to ? "mailto:" + to + "?subject=" + encodeURIComponent(subject) + "&body=" + encodeURIComponent(body) : ""; }
function userByName(data, username) { return (data?.users || []).find((u) => u.username === username) || null; }
function magTargets(data) { return (data?.users || []).filter((u) => u.role === "magasinier" || u.role === "admin"); }
function changeSummary(o) { const p = o?.prep || {}; let miss = 0, repl = 0; Object.values(p).forEach((x) => { if (x?.state === "manquant") miss++; else if (x?.state === "remplace" || x?.repl) repl++; }); const parts = []; if (repl) parts.push(repl + " remplacement(s)"); if (miss) parts.push(miss + " manquant(s)"); return parts.length ? "\nModifications : " + parts.join(", ") + "." : ""; }
function msgNewDemande(o) { return "PomProc — Nouvelle demande " + o.ref + "\nChantier : " + o.nom + " (N° " + o.numero + ")\nDépart : " + fmtD(o.departDate) + (o.departTime ? " à " + o.departTime : "") + "\nDemandeur : " + o.username + "\n" + (o.lines?.length || 0) + " article(s) à préparer."; }
function msgDone(o) { return "PomProc — Demande " + o.ref + " prête \u2705\nChantier : " + o.nom + " (N° " + o.numero + ")\nPréparée par : " + (o.completedBy || "—") + "\n" + (o.lines?.length || 0) + " article(s)." + changeSummary(o); }

function NotifyActions({ targets, subject, body }) {
  const emails = (targets || []).map((t) => t.email).filter(Boolean);
  const withPhone = (targets || []).filter((t) => onlyDigits(t.phone));
  if (!emails.length && !withPhone.length) return <div className="notify-empty">Aucun contact enregistré. Ajoutez e-mail / téléphone dans <b>Admin ▸ Utilisateurs</b>.</div>;
  const waText = subject + "\n\n" + body;
  return (
    <div className="notify-row">
      {emails.length > 0 && <a className="ntf-btn mail" href={mailLink(emails, subject, body)}>✉︎ E-mail</a>}
      {withPhone.map((t, i) => <a key={i} className="ntf-btn wa" href={waLink(t.phone, waText)} target="_blank" rel="noreferrer">✆ WhatsApp{withPhone.length > 1 ? " · " + t.username : ""}</a>)}
    </div>
  );
}
async function prepImageUpload(file) { return await fileToThumb(file, 1000, 0.72); }
function fileToThumb(file, max = 320, quality = 0.82) {
  return new Promise((resolve, reject) => {
    const img = new Image(); const url = URL.createObjectURL(file);
    img.onload = () => {
      const scale = Math.min(1, max / Math.max(img.width, img.height));
      const w = Math.round(img.width * scale), h = Math.round(img.height * scale);
      const c = document.createElement("canvas"); c.width = w; c.height = h;
      c.getContext("2d").drawImage(img, 0, 0, w, h); URL.revokeObjectURL(url);
      try { resolve(c.toDataURL("image/jpeg", quality)); } catch (e) { reject(e); }
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error("Image illisible")); };
    img.src = url;
  });
}

/* ===== Icônes catégories (SVG inline, robustes) ===== */
const ICONS = {
  epi: <><path d="M3 18h18" /><path d="M5 18a7 7 0 0 1 14 0" /><path d="M10 6h4v5" /></>,
  conso: <><path d="M9 3h6v3l1 2v11a2 2 0 0 1-2 2h-4a2 2 0 0 1-2-2V8l1-2z" /><path d="M9 12h6" /></>,
  mat: <><path d="M15 4l5 5" /><path d="M13 6l5 5" /><path d="M4 20l8-8 2 2-8 8z" /><path d="M4 20v-3h3" /></>,
  out: <><path d="M15 7a4 4 0 0 1-5 5l-5 5 3 3 5-5a4 4 0 0 1 5-5l-2 2-2-2 2-2z" /></>,
  mac: <><circle cx="12" cy="12" r="3.2" /><path d="M12 2v3M12 19v3M2 12h3M19 12h3M5 5l2 2M17 17l2 2M19 5l-2 2M7 17l-2 2" /></>,
  carb: <><path d="M4 8h11v11H4z" /><path d="M15 10h3v6h-3" /><path d="M6 8V6h4v2" /><path d="M8 12l2 3h-3z" /></>,
  pou: <><path d="M4 7h16" /><path d="M6 7l1 13h10l1-13" /><path d="M9 7V4h6v3" /><path d="M10 11v6M14 11v6" /></>,
  vit: <><rect x="4" y="4" width="16" height="16" rx="1" /><path d="M4 10h16M10 4v16" /></>,
  ech: <><path d="M7 3v18M17 3v18M7 7h10M7 11h10M7 15h10" /></>,
  epc: <><path d="M10 4h4l4 15H6z" /><path d="M9 10h6M8 14h8" /><path d="M4 20h16" /></>,
  tex: <><path d="M8 3L4 6l2 3 2-1v10h8V8l2 1 2-3-4-3-2 2-2-2z" /></>,
};
function CatIcon({ icon, color, size = 18 }) {
  return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color || "currentColor"} strokeWidth="1.9" strokeLinecap="round" strokeLinejoin="round">{ICONS[icon] || ICONS.mat}</svg>;
}

/* ==================================================================== */
export default function App() {
  const [ready, setReady] = useState(false);
  const [data, setData] = useState(null);
  const [orders, setOrders] = useState([]);
  const [session, setSession] = useState(null);
  const [view, setView] = useState("login");
  const [activeOrder, setActiveOrder] = useState(null);

  useEffect(() => {
    (async () => {
      let d = await sget(KEY_DATA);
      if (!d) {
        d = { users: [{ username: "admin", password: "admin", workplace: "IIS", role: "admin" }, { username: "magasin", password: "magasin", workplace: "Stock IIS", role: "magasinier" }], productCats: SEED.productCategories, products: SEED.products, templates: SEED.templates, counters: {}, kits: (SEED.kits || []) };
        await sset(KEY_DATA, d);
      }
      d.counters = d.counters || {};
      let o = await sget(KEY_ORDERS); if (!o) { o = []; await sset(KEY_ORDERS, o); }
      setData(d); setOrders(o); setReady(true);
    })();
  }, []);

  const saveData = useCallback(async (n) => { setData(n); await sset(KEY_DATA, n); }, []);
  const saveOrders = useCallback(async (n) => { setOrders(n); await sset(KEY_ORDERS, n); }, []);
  const updateOrder = useCallback(async (o) => { setOrders((prev) => { const n = prev.map((x) => (x.id === o.id ? o : x)); sset(KEY_ORDERS, n); return n; }); setActiveOrder(o); }, []);

  const productById = useMemo(() => { const m = {}; (data?.products || []).forEach((p) => (m[p.id] = p)); return m; }, [data]);
  const addToOrder = useCallback((order, payload) => {
    const now = Date.now();
    const pForId = (id) => productById[id] || (order.extras || []).find((e) => e.id === id) || (payload.extras || []).find((e) => e.id === id);
    const newLines = (payload.lines || []).map((l) => ({ ...l, addedAt: now, addedBy: session.username, addedRole: session.role }));
    if (!newLines.length) return;
    let addFFA = 0, addDC = 0, addTotal = 0;
    newLines.forEach((l) => { const p = pForId(l.productId); const t = lineTotal(l, p); addTotal += t; if (l.entity === "FFA") addFFA += t; else if (l.entity === "DC") addDC += t; });
    const merged = { ...order, lines: [...order.lines, ...newLines], extras: [...(order.extras || []), ...(payload.extras || [])], total: (order.total || 0) + addTotal, billFFA0: (order.billFFA0 ?? order.billFFA ?? 0) + addFFA, billDC0: (order.billDC0 ?? order.billDC ?? 0) + addDC, billFFA: (order.billFFA ?? 0) + addFFA, billDC: (order.billDC ?? 0) + addDC, status: order.status === "terminee" ? "en_cours" : order.status, amendedAt: now };
    updateOrder(merged);
  }, [productById, session, updateOrder]);
  const catById = useMemo(() => { const m = {}; (data?.productCats || []).forEach((c) => (m[c.id] = c)); return m; }, [data]);

  function logout() { setSession(null); setView("login"); }
  const open = (o) => { setActiveOrder(o); setView("sheet"); };

  let s = null;
  if (!ready) s = <Splash />;
  else if (!session && view === "login") s = <Login data={data} onLogin={(u) => { setSession(u); setView("home"); }} goRegister={() => setView("register")} />;
  else if (!session && view === "register") s = <Register data={data} saveData={saveData} onDone={(u) => { setSession(u); setView("home"); }} goBack={() => setView("login")} />;
  else if (view === "home") s = <Home session={session} orders={orders} data={data} onNew={() => setView("order")} onOpen={open} nav={setView} logout={logout} />;
  else if (view === "order") s = <OrderWizard session={session} data={data} saveData={saveData} productById={productById} catById={catById} orders={orders} onCancel={() => setView("home")} onSaved={(o) => { saveOrders([o, ...orders]); open(o); }} />;
  else if (view === "additems") s = <OrderWizard session={session} data={data} saveData={saveData} productById={productById} catById={catById} orders={orders} addMode baseOrder={activeOrder} onCancel={() => setView("sheet")} onAdd={(payload) => { addToOrder(activeOrder, payload); setView("sheet"); }} />;
  else if (view === "sheet") s = <Sheet order={activeOrder} session={session} data={data} productById={productById} catById={catById} addItems={() => setView("additems")} back={() => setView("home")} />;
  else if (view === "templates") s = <Templates session={session} data={data} saveData={saveData} productById={productById} catById={catById} back={() => setView("home")} />;
  else if (view === "magasin") s = <Magasin orders={orders} onOpen={(o) => { setActiveOrder(o); setView("prepare"); }} back={() => setView("home")} />;
  else if (view === "prepare") s = <Prepare order={activeOrder} session={session} data={data} productById={productById} catById={catById} updateOrder={updateOrder} addItems={() => setView("additems")} back={() => setView("magasin")} viewSheet={() => setView("sheet")} />;
  else if (view === "allorders") s = <AllOrders orders={orders} onOpen={open} back={() => setView("home")} />;
  else if (view === "billing") s = <Billing orders={orders} productById={productById} session={session} updateOrder={updateOrder} onOpen={open} back={() => setView("home")} />;
  else if (view === "finished") s = <Finished orders={orders} onOpen={open} back={() => setView("home")} />;
  else if (view === "analyse") s = <Analytics orders={orders} data={data} productById={productById} catById={catById} back={() => setView("home")} />;
  else if (view === "admin") s = <Admin data={data} saveData={saveData} back={() => setView("home")} />;

  return (<><style>{CSS}</style><div className="app">{s}</div></>);
}

/* ============================ UI de base ============================ */
function Splash() { return <div className="center"><div className="spinner" /></div>; }
function EntBadge({ e }) { const m = ENTITIES[e] || ENTITIES.IIS; return <span className="ebadge" style={{ background: m.color }}>{m.label}</span>; }
function StatusBadge({ st }) { const m = STATUS[st] || STATUS.envoyee; return <span className="stbadge" style={{ background: m.color }}>{m.label}</span>; }
function CatTag({ cat, size }) { return <span className="cattag" style={{ background: cat.color }}><CatIcon icon={cat.icon} color="#fff" size={size || 13} /></span>; }

function TopBar({ title, right, onBack }) {
  return (
    <div className="topbar no-print">
      {onBack ? <button className="icon-btn" onClick={onBack} aria-label="Retour">‹</button> : <img src={LOGO} alt="XLG" className="topbar-logo" />}
      <div className="topbar-title">{title}</div>
      <div className="topbar-right">{right}</div>
    </div>
  );
}

function Login({ data, onLogin, goRegister }) {
  const [u, setU] = useState(""); const [p, setP] = useState(""); const [err, setErr] = useState("");
  function submit() { const f = (data.users || []).find((x) => x.username.toLowerCase() === u.trim().toLowerCase() && x.password === p); if (!f) return setErr("Identifiant ou mot de passe incorrect."); onLogin(f); }
  return (
    <div className="auth">
      <div className="auth-hero"><img src={LOGO} alt="XLG" className="auth-logo" /><div className="auth-app">POM<span>PROC</span></div><div className="auth-iis">by <b>IIS</b></div></div>
      <div className="card auth-card">
        <label className="lbl">Identifiant</label><input className="inp" value={u} onChange={(e) => setU(e.target.value)} placeholder="prénom.nom" autoCapitalize="none" />
        <label className="lbl">Mot de passe</label><input className="inp" type="password" value={p} onChange={(e) => setP(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} placeholder="••••••" />
        {err && <div className="err">{err}</div>}
        <button className="btn btn-primary btn-block" onClick={submit}>Se connecter</button>
        <button className="btn btn-ghost btn-block" onClick={goRegister}>Créer un compte</button>
      </div>
      <div className="auth-note">Démo — admin : <b>admin</b>/<b>admin</b> · magasinier : <b>magasin</b>/<b>magasin</b></div>
    </div>
  );
}

function Register({ data, saveData, onDone, goBack }) {
  const [f, setF] = useState({ username: "", password: "", workplace: "", email: "", phone: "" }); const [err, setErr] = useState("");
  async function submit() {
    const un = f.username.trim();
    if (!un || !f.password || !f.workplace.trim()) return setErr("Identifiant, mot de passe et lieu de travail sont requis.");
    if (data.users.some((x) => x.username.toLowerCase() === un.toLowerCase())) return setErr("Cet identifiant existe déjà.");
    const user = { username: un, password: f.password, workplace: f.workplace.trim(), role: "user", email: f.email.trim(), phone: f.phone.trim() };
    await saveData({ ...data, users: [...data.users, user] }); onDone(user);
  }
  return (
    <div className="auth"><TopBar title="Créer un compte" onBack={goBack} />
      <div className="card auth-card">
        <label className="lbl">Identifiant</label><input className="inp" value={f.username} onChange={(e) => setF({ ...f, username: e.target.value })} placeholder="prénom.nom" autoCapitalize="none" />
        <label className="lbl">Mot de passe</label><input className="inp" type="password" value={f.password} onChange={(e) => setF({ ...f, password: e.target.value })} />
        <label className="lbl">Lieu de travail</label><input className="inp" value={f.workplace} onChange={(e) => setF({ ...f, workplace: e.target.value })} placeholder="Agence / site" />
        <label className="lbl">E-mail <span className="lbl-opt">(pour les notifications)</span></label><input className="inp" type="email" value={f.email} onChange={(e) => setF({ ...f, email: e.target.value })} placeholder="prenom.nom@xlg.be" autoCapitalize="none" inputMode="email" />
        <label className="lbl">Téléphone WhatsApp <span className="lbl-opt">(format international, ex : 32470…)</span></label><input className="inp" type="tel" value={f.phone} onChange={(e) => setF({ ...f, phone: e.target.value })} placeholder="32 470 12 34 56" inputMode="tel" />
        {err && <div className="err">{err}</div>}
        <button className="btn btn-primary btn-block" onClick={submit}>Créer mon compte</button>
      </div>
    </div>
  );
}

function Home({ session, orders, data, onNew, onOpen, nav, logout }) {
  const mine = orders.filter((o) => o.username === session.username);
  const done = mine.filter((o) => o.status === "terminee");
  const isStock = session.role === "magasinier" || session.role === "admin";
  const admin = session.role === "admin";
  const toPrepare = orders.filter((o) => o.status !== "terminee").length;
  const finishedCount = orders.filter((o) => o.status === "terminee").length;
  const tplCount = (data.templates || []).length;
  return (
    <div>
      <TopBar title="Accueil" right={<button className="pill pill-out" onClick={logout}>Quitter</button>} />
      <div className="scroll">
        <div className="hello"><div className="hello-name">Bonjour {session.username}</div><div className="hello-sub">{session.workplace} · {ROLES.find((r) => r.id === session.role)?.label}</div></div>
        {done.length > 0 && <div className="notif" onClick={() => onOpen(done[0])}>✓ {done.length} préparation(s) terminée(s) — appuyez pour voir le compte-rendu</div>}
        <button className="cta" onClick={onNew}><span className="cta-plus">+</span><span><b>Nouvelle demande</b><small>Préparer une commande chantier</small></span></button>
        <div className="quick">
          <button className="qbtn" onClick={() => nav("templates")}><span className="qi">▦</span>Templates{tplCount ? <em className="qi-badge">{tplCount}</em> : null}</button>
          {isStock && <button className="qbtn" onClick={() => nav("magasin")}><span className="qi">⧉</span>Magasin{toPrepare ? <em className="qi-badge">{toPrepare}</em> : null}</button>}
          {admin && <button className="qbtn" onClick={() => nav("finished")}><span className="qi">✓</span>Terminées{finishedCount ? <em className="qi-badge">{finishedCount}</em> : null}</button>}
          {admin && <button className="qbtn" onClick={() => nav("allorders")}><span className="qi">≣</span>Toutes les demandes{orders.length ? <em className="qi-badge">{orders.length}</em> : null}</button>}
          {admin && <button className="qbtn" onClick={() => nav("billing")}><span className="qi">€</span>Facturation</button>}
          {admin && <button className="qbtn" onClick={() => nav("analyse")}><span className="qi">📊</span>Analyse</button>}
          {admin && <button className="qbtn" onClick={() => nav("admin")}><span className="qi">⚙</span>Administration</button>}
        </div>
        <div className="sec-h">Mes demandes</div>
        {mine.length === 0 ? <div className="empty">Aucune demande pour le moment.</div> :
          <div className="list">{mine.map((o) => <OrderRow key={o.id} o={o} onClick={() => onOpen(o)} />)}</div>}
        <div className="version">PomProc {APP_VERSION}</div>
      </div>
    </div>
  );
}function OrderRow({ o, onClick, showUser }) {
  return (
    <button className="row" onClick={onClick}>
      <div className="row-main">
        <div className="row-title">{o.nom} <span className="ref-tag">{o.ref}</span></div>
        <div className="row-sub">{o.lines.length} art. · {showUser ? o.username + " · " : ""}départ {fmtD(o.departDate)}{o.departTime ? " " + o.departTime : ""}</div>
      </div>
      <div className="row-right"><StatusBadge st={o.status} /><span className="chev">›</span></div>
    </button>
  );
}

/* ======================= Assistant de demande ======================= */
function OrderWizard({ session, data, saveData, productById, catById, orders, onCancel, onSaved, addMode, baseOrder, onAdd }) {
  const [step, setStep] = useState(addMode ? 2 : 1);
  const [draft, setDraft] = useState({ nom: "", numero: "", responsable: "", departDate: "", departTime: "", templateName: "", lines: [], extras: [] });
  const cats = useMemo(() => [...data.productCats].sort((a, b) => a.order - b.order), [data.productCats]);
  const [chip, setChip] = useState(cats[0]?.id || "all");
  const [q, setQ] = useState("");
  const [vsel, setVsel] = useState({});
  const [custOpen, setCustOpen] = useState(false);
  const [detailsFor, setDetailsFor] = useState(null);
  const [basketOpen, setBasketOpen] = useState(false);
  const [cf, setCf] = useState({ nom: "", ref: "", unite: "Pièce", prix: "", qty: "1", catId: cats[1]?.id || cats[0]?.id || "", entity: "FFA" });
  const pById = useMemo(() => { const m = { ...productById }; (draft.extras || []).forEach((e) => (m[e.id] = e)); return m; }, [productById, draft.extras]);
  const selFor = (p, groups) => { const cur = vsel[p.id] || {}; const m = {}; groups.forEach((g) => (m[g.name] = cur[g.name] || g.options[0])); return m; };
  const comboLabel = (groups, sel) => groups.map((g) => sel[g.name]).join(" / ");
  const keyOf = (pid, v) => (v ? pid + "::" + v : pid);
  const lineFor = (pid, v) => draft.lines.find((l) => lineKey(l) === keyOf(pid, v));
  const qtyOf = (pid) => draft.lines.filter((l) => l.productId === pid).reduce((s, l) => s + (l.qty || 0), 0);
  const total = useMemo(() => draft.lines.reduce((s, l) => s + lineTotal(l, pById[l.productId]), 0), [draft.lines, pById]);
  const freqList = useMemo(() => { const tally = {}; (orders || []).forEach((o) => (o.lines || []).forEach((l) => { tally[l.productId] = (tally[l.productId] || 0) + (l.qty || 1); })); return Object.entries(tally).sort((a, b) => b[1] - a[1]).map(([id]) => productById[id]).filter(Boolean).slice(0, 24); }, [orders, productById]);
  const pastOrders = useMemo(() => [...(orders || [])].sort((a, b) => b.createdAt - a.createdAt).slice(0, 20), [orders]);
  function setLine(pid, patch, v) {
    setDraft((d) => { const k = keyOf(pid, v); const i = d.lines.findIndex((l) => lineKey(l) === k); let lines;
      if (i === -1) { const p = pById[pid]; lines = [...d.lines, { productId: pid, variant: v || "", qty: 1, mode: "unite", entity: p?.entity || "IIS", assetNo: "", ...patch }]; }
      else { lines = d.lines.slice(); lines[i] = { ...lines[i], ...patch }; }
      return { ...d, lines: lines.filter((l) => l.qty > 0) }; });
  }
  const removeKey = (k) => setDraft((d) => ({ ...d, lines: d.lines.filter((l) => lineKey(l) !== k) }));
  const setProdEntity = (pid, e) => setDraft((d) => ({ ...d, lines: d.lines.map((l) => (l.productId === pid ? { ...l, entity: e } : l)) }));
  const setLineEntity = (k, e) => setDraft((d) => ({ ...d, lines: d.lines.map((l) => (lineKey(l) === k ? { ...l, entity: e } : l)) }));
  function addCustom() {
    if (!cf.nom.trim()) return;
    const id = "cust_" + uid("");
    const e = { id, custom: true, catId: cf.catId, ref: cf.ref.trim() || "—", nom: cf.nom.trim(), unite: cf.unite.trim() || "Pièce", qteEmballage: 1, prix: parseFloat(String(cf.prix).replace(",", ".")) || 0, entity: cf.entity };
    const qn = parseInt(cf.qty, 10) || 1;
    setDraft((d) => ({ ...d, extras: [...(d.extras || []), e], lines: [...d.lines, { productId: id, variant: "", qty: qn, mode: "unite", entity: cf.entity, assetNo: "" }] }));
    setCf({ nom: "", ref: "", unite: "Pièce", prix: "", qty: "1", catId: cf.catId, entity: cf.entity }); setCustOpen(false);
  }
  function applyTemplate(t) { if (!t) { setDraft((d) => ({ ...d, templateName: "" })); return; } const lines = t.lines.map((l) => { const p = pById[l.productId]; return p ? { productId: l.productId, variant: l.variant || "", qty: l.qty, mode: l.mode || "unite", entity: p.entity || "IIS", assetNo: "" } : null; }).filter(Boolean); setDraft((d) => ({ ...d, templateName: t.name, lines })); }
  function applyPastOrder(o) { if (!o) return; const extras = (o.extras || []).map((e) => ({ ...e })); const extraIds = new Set(extras.map((e) => e.id)); const lines = (o.lines || []).filter((l) => productById[l.productId] || extraIds.has(l.productId)).map((l) => ({ productId: l.productId, variant: l.variant || "", qty: l.qty, mode: l.mode || "unite", entity: l.entity || "IIS", assetNo: "" })); setDraft((d) => ({ ...d, extras, lines, templateName: "" })); setStep(2); }
  function applyKit(kit) { if (!kit) return; setDraft((d) => { const lines = [...d.lines]; (kit.items || []).forEach((it) => { const p = pById[it.productId]; if (!p) return; const variant = it.variant || ""; const k = variant ? it.productId + "::" + variant : it.productId; const idx = lines.findIndex((l) => lineKey(l) === k); if (idx >= 0) lines[idx] = { ...lines[idx], qty: lines[idx].qty + (it.qty || 1) }; else lines.push({ productId: it.productId, variant, qty: it.qty || 1, mode: it.mode || "unite", entity: it.entity || p.entity || "IIS", assetNo: "" }); }); return { ...d, lines }; }); }
  async function saveAsTemplate() { const name = window.prompt("Nom du template :"); if (!name) return; const t = { id: uid("tpl_"), name: name.trim(), ownerUsername: session.username, shared: true, lines: draft.lines.map((l) => ({ productId: l.productId, variant: l.variant || "", qty: l.qty, mode: l.mode })), createdAt: Date.now() }; await saveData({ ...data, templates: [...(data.templates || []), t] }); window.alert("Template enregistré."); }
  const canNext = step === 1 ? draft.nom.trim() && draft.numero.trim() && draft.departDate : step === 2 ? draft.lines.length > 0 : true;
  const shown = useMemo(() => {
    if (q.trim()) { const s = q.trim().toLowerCase(); return [...data.products].filter((p) => p.nom.toLowerCase().includes(s) || String(p.ref).toLowerCase().includes(s)).sort(byName); }
    if (chip === "freq") return freqList;
    if (chip === "all") return [...data.products].sort(byName);
    return [...data.products].filter((p) => p.catId === chip).sort(byName);
  }, [q, chip, data.products, freqList]);
  async function finish() {
    if (addMode) { onAdd({ lines: draft.lines, extras: draft.extras }); return; }
    const now = Date.now(); const ym = new Date(now).toISOString().slice(2, 7).replace("-", ""); const next = (data.counters?.[ym] || 0) + 1; const ref = ym + "-" + String(next).padStart(3, "0");
    await saveData({ ...data, counters: { ...(data.counters || {}), [ym]: next } });
    let billFFA = 0, billDC = 0; draft.lines.forEach((l) => { const t = lineTotal(l, pById[l.productId]); if (l.entity === "FFA") billFFA += t; else if (l.entity === "DC") billDC += t; });
    onSaved({ id: uid("o_"), ref, username: session.username, workplace: session.workplace, nom: draft.nom.trim(), numero: draft.numero.trim(), responsable: draft.responsable.trim(), departDate: draft.departDate, departTime: draft.departTime, templateName: draft.templateName, lines: draft.lines, extras: draft.extras, total, billFFA, billDC, billFFA0: billFFA, billDC0: billDC, createdAt: now, status: "envoyee", assignedTo: "", completedBy: "", completedAt: 0, prep: {}, magasinierNote: "", paymentStatus: "non_paye" });
  }
  const renderProd = (p) => {
    const cat = catById[p.catId]; const groups = variantGroupsOf(p, catById); const hasVar = groups.length > 0;
    const on = draft.lines.some((l) => l.productId === p.id); const tot = qtyOf(p.id); const rowEnt = draft.lines.find((l) => l.productId === p.id)?.entity || p.entity;
    return (
      <div key={p.id} className={"prod" + (on ? " prod-on" : "")} style={on ? { borderColor: cat?.color } : {}}>
        <div className="prod-top">
          <div className="prod-thumb" style={{ borderColor: cat?.color }}>{p.image ? <img src={p.image} alt="" /> : <CatIcon icon={cat?.icon} color={cat?.color} size={20} />}{on && tot > 0 ? <span className="qty-badge" style={{ background: cat?.color }}>{tot}</span> : null}</div>
          <div className="prod-info"><div className="prod-name">{p.nom} <EntBadge e={rowEnt} /></div><div className="prod-meta">Réf {p.ref} · {p.unite}{p.prix ? " · " + money(p.prix) : ""}{p.qteEmballage > 1 ? " · emb. " + p.qteEmballage : ""}{q ? " · " + (cat?.name || "") : ""}</div></div>
          {!hasVar && <Stepper value={lineFor(p.id)?.qty || 0} onChange={(v) => setLine(p.id, { qty: v })} />}
        </div>
        {hasVar && groups.length === 1 && <div className="var-grid">{groups[0].options.map((sz) => { const l = lineFor(p.id, sz); return <div key={sz} className={"var-cell" + (l ? " on" : "")}><span className="var-lbl">{sz}</span><Stepper value={l?.qty || 0} onChange={(v) => setLine(p.id, { qty: v }, sz)} /></div>; })}</div>}
        {hasVar && groups.length >= 2 && (() => { const sel = selFor(p, groups); const label = comboLabel(groups, sel); const combo = lineFor(p.id, label); return (
          <div className="var-multi">
            {groups.map((g) => <div key={g.name} className="vg"><span className="vg-name">{g.name}</span><div className="vg-opts">{g.options.map((opt) => <button key={opt} className={"vg-opt" + (sel[g.name] === opt ? " on" : "")} onClick={() => setVsel((v) => ({ ...v, [p.id]: { ...selFor(p, groups), ...(v[p.id] || {}), [g.name]: opt } }))}>{opt}</button>)}</div></div>)}
            <div className="vg-add"><span className="vg-combo">{label}</span><Stepper value={combo?.qty || 0} onChange={(v) => setLine(p.id, { qty: v }, label)} /></div>
          </div>); })()}
        {hasVar && draft.lines.some((l) => l.productId === p.id) && <div className="vg-list">{draft.lines.filter((l) => l.productId === p.id).sort((a, b) => (a.variant || "").localeCompare(b.variant || "")).map((l) => <div key={lineKey(l)} className="vg-chip2"><b>{l.variant || "—"}</b> ×{l.qty}<div className="vg-ent">{ENT_ORDER.map((e) => <button key={e} className={"vg-eb" + (l.entity === e ? " on" : "")} style={l.entity === e ? { background: ENTITIES[e].color, borderColor: ENTITIES[e].color, color: "#fff" } : { color: ENTITIES[e].color }} onClick={() => setLineEntity(lineKey(l), e)}>{e}</button>)}</div><button className="vg-x" onClick={() => removeKey(lineKey(l))}>✕</button></div>)}</div>}
        {on && (
          <div className="row-opts">
            <button className="details-btn" onClick={() => setDetailsFor(p.id)}>⚙ Détails{!hasVar && lineFor(p.id)?.mode === "paquet" ? " · emb." : ""}{!hasVar && cat?.num && lineFor(p.id)?.assetNo ? " · " + lineFor(p.id).assetNo : ""}</button>
            {(() => { const chain = assocClosure(p.id, (id) => pById[id]); return chain.length > 0 ? (<div className="assoc inline"><span>À prévoir aussi :</span>{chain.map((aid) => { const ap = pById[aid]; if (!ap) return null; const has = draft.lines.some((l) => l.productId === aid); return <button key={aid} className={"assoc-chip" + (has ? " has" : "")} onClick={() => !has && setLine(aid, { qty: 1 })}>{has ? "✓ " : "+ "}{ap.nom}</button>; })}</div>) : null; })()}
          </div>
        )}
      </div>);
  };
  return (
    <div>
      <TopBar title={addMode ? "Ajouter à " + baseOrder.ref : "Nouvelle demande"} onBack={(addMode ? step <= 2 : step === 1) ? onCancel : () => setStep(step - 1)} right={addMode ? null : <div className="stepdots">{[1, 2, 3].map((n) => <span key={n} className={"dot" + (n === step ? " on" : n < step ? " done" : "")} />)}</div>} />
      <div className="scroll pb-bar">
        {step === 1 && (
          <div className="pane">
            <h2 className="pane-h">Informations</h2>
            <label className="lbl">Nom du chantier</label><input className="inp" value={draft.nom} onChange={(e) => setDraft({ ...draft, nom: e.target.value })} placeholder="Ex : Résidence Louise" />
            <label className="lbl">Numéro de chantier</label><input className="inp" value={draft.numero} onChange={(e) => setDraft({ ...draft, numero: e.target.value })} placeholder="Ex : 10432" inputMode="numeric" />
            <div className="row2b"><div><label className="lbl">Date de départ</label><input className="inp" type="date" value={draft.departDate} onChange={(e) => setDraft({ ...draft, departDate: e.target.value })} /></div><div><label className="lbl">Heure de départ</label><input className="inp" type="time" value={draft.departTime} onChange={(e) => setDraft({ ...draft, departTime: e.target.value })} /></div></div>
            <label className="lbl">Chef d'équipe (facultatif)</label><input className="inp" value={draft.responsable} onChange={(e) => setDraft({ ...draft, responsable: e.target.value })} placeholder="Responsable chantier" />
            <label className="lbl">Partir d'un template (facultatif)</label>
            <select className="inp" value={draft.templateName} onChange={(e) => applyTemplate((data.templates || []).find((t) => t.name === e.target.value))}><option value="">— Aucun —</option>{(data.templates || []).map((t) => <option key={t.id} value={t.name}>{t.name} ({t.lines.length})</option>)}</select>
            {pastOrders.length > 0 && (<><label className="lbl">Repartir d'une demande passée (facultatif)</label>
            <select className="inp" value="" onChange={(e) => applyPastOrder(pastOrders.find((x) => x.id === e.target.value))}><option value="">— Choisir une demande —</option>{pastOrders.map((o) => <option key={o.id} value={o.id}>{o.ref} — {o.nom} ({o.lines.length} art.)</option>)}</select></>)}
            <div className="demandeur-note">Demandeur : <b>{session.username}</b> · une référence unique sera attribuée</div>
          </div>
        )}
        {step === 2 && (
          <div className="pane">
            {addMode && <div className="add-banner">Ajout à <b>{baseOrder.ref}</b> — {baseOrder.nom}. Les {baseOrder.lines.length} article(s) déjà présents sont conservés ; ceux-ci seront ajoutés et datés d'aujourd'hui.</div>}
            <div className="prod-head2"><h2 className="pane-h">Produits</h2><button className="mini-btn hc-btn" onClick={() => setCustOpen(true)}>+ Hors cat.</button></div>
            <div className="filterbar">
              <div className="search"><span>⌕</span><input placeholder="Rechercher un produit ou une réf…" value={q} onChange={(e) => setQ(e.target.value)} />{q && <button onClick={() => setQ("")}>✕</button>}</div>
              {!q && (<div className="chips">
                {freqList.length > 0 && <button className={"chip" + (chip === "freq" ? " on" : "")} onClick={() => setChip("freq")}>★ Fréquents</button>}
                {(data.kits || []).length > 0 && <button className={"chip" + (chip === "kits" ? " on" : "")} onClick={() => setChip("kits")}>🎁 Kits</button>}
                <button className={"chip" + (chip === "all" ? " on" : "")} onClick={() => setChip("all")}>Tous</button>
                {cats.map((c) => { const n = draft.lines.filter((l) => pById[l.productId]?.catId === c.id).length; return <button key={c.id} className={"chip" + (chip === c.id ? " on" : "")} style={chip === c.id ? { background: c.color, borderColor: c.color, color: "#fff" } : { borderColor: c.color }} onClick={() => setChip(c.id)}>{c.name}{n > 0 ? " · " + n : ""}</button>; })}
              </div>)}
            </div>
            {(draft.extras || []).length > 0 && (
              <div className="extras-box"><div className="extras-h">Produits hors catalogue</div>{draft.extras.map((e) => { const l = lineFor(e.id); return (
                <div key={e.id} className="extra-row"><span className="extra-name">{e.nom} <EntBadge e={e.entity} /></span><span className="extra-meta">{money(e.prix)}</span><Stepper value={l?.qty || 0} onChange={(v) => setLine(e.id, { qty: v })} /></div>); })}</div>)}
            <div className="prodlist">
              {chip === "kits" && !q ? (
                (data.kits || []).length === 0 ? <div className="empty sm">Aucun kit défini.</div> :
                (data.kits || []).map((kit) => (
                  <div key={kit.id} className="kit-card">
                    <div className="kit-card-h"><span className="kit-card-name">🎁 {kit.name}</span><button className="kit-add" onClick={() => applyKit(kit)}>+ Ajouter</button></div>
                    <div className="kit-card-items">{kit.items.map((it, i) => { const p = pById[it.productId]; return <span key={i} className="kit-chip">{it.qty}× {p?.nom || it.productId}{it.variant ? " (" + it.variant + ")" : ""}</span>; })}</div>
                  </div>))
              ) : (<>
                {chip === "freq" && !q && freqList.length === 0 && <div className="empty sm">Pas encore d'historique — choisissez une catégorie.</div>}
                {shown.map((p) => renderProd(p))}
                {shown.length === 0 && q && <div className="empty sm">Aucun produit trouvé.</div>}
              </>)}
            </div>
          </div>
        )}
        {step === 3 && (
          <div className="pane">
            <h2 className="pane-h">{addMode ? "Confirmer l'ajout" : "Récapitulatif"}</h2>
            {addMode
              ? <div className="recap-box"><div className="recap-line"><span>Ajout à</span><b>{baseOrder.ref} — {baseOrder.nom}</b></div><div className="recap-line"><span>Ajouté par</span><b>{session.username}</b></div><div className="recap-line"><span>Date</span><b>{fmtD(new Date().toISOString().slice(0, 10))}</b></div></div>
              : <div className="recap-box">
              <div className="recap-line"><span>Chantier</span><b>{draft.nom} — N° {draft.numero}</b></div>
              <div className="recap-line"><span>Départ</span><b>{fmtD(draft.departDate)}{draft.departTime ? " à " + draft.departTime : ""}</b></div>
              {draft.templateName && <div className="recap-line"><span>Template</span><b>{draft.templateName}</b></div>}
              <div className="recap-line"><span>Demandeur</span><b>{session.username}</b></div>
            </div>}
            <div className="recap-items">
              {[...draft.lines].sort((a, b) => byName(pById[a.productId] || { nom: "" }, pById[b.productId] || { nom: "" })).map((l) => { const p = pById[l.productId]; const cat = catById[p?.catId]; return (
                <div key={lineKey(l)} className="ritem">
                  <CatTag cat={cat || { color: "#999", icon: "mat" }} />
                  <div className="ritem-main"><div className="ritem-name">{p?.nom}{l.variant ? " · " + l.variant : ""} <EntBadge e={l.entity} />{p?.custom ? <span className="asset-tag">hors cat.</span> : null}{l.assetNo ? <span className="asset-tag">{l.assetNo}</span> : null}</div><div className="ritem-sub">{l.qty} × {l.mode === "paquet" ? "emb. " + p?.qteEmballage : (p?.unite || "").toLowerCase()} · Réf {p?.ref}</div></div>
                  <div className="ritem-amt">{money(lineTotal(l, p))}<button className="del" onClick={() => removeKey(lineKey(l))}>✕</button></div>
                </div>); })}
            </div>
            {total > 0 && <div className="recap-total"><span>Total indicatif</span><b>{money(total)}</b></div>}
            {!addMode && <button className="btn btn-ghost btn-block" onClick={saveAsTemplate}>▦ Enregistrer comme template</button>}
          </div>
        )}
      </div>
      <div className="bottombar no-print">
        {step === 2
          ? (<><button className="basket-btn" onClick={() => setBasketOpen(true)}>🧺 Panier<span className="bb-count">{draft.lines.length}</span></button><div className="bb-total">{total > 0 ? <b>{money(total)}</b> : <b>—</b>}</div><button className="btn btn-primary" disabled={!canNext} onClick={() => setStep(step + 1)}>Continuer</button></>)
          : (<><div className="bb-total" />{step < 3 ? <button className="btn btn-primary" disabled={!canNext} onClick={() => setStep(step + 1)}>Continuer</button> : <button className="btn btn-primary" onClick={finish}>{addMode ? "Ajouter à la demande" : "Envoyer au magasin"}</button>}</>)}
      </div>
      {detailsFor && (() => {
        const p = pById[detailsFor]; if (!p) return null;
        const cat = catById[p.catId]; const groups = variantGroupsOf(p, catById); const hasVar = groups.length > 0;
        const line = draft.lines.find((l) => l.productId === p.id); const curEnt = line?.entity || p.entity;
        return (
          <div className="modal-bg" onClick={() => setDetailsFor(null)}><div className="modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-h">{p.nom}</div>
            {!hasVar && (<><label className="lbl">Conditionnement</label><div className="seg"><button className={"seg-b" + ((line?.mode || "unite") === "unite" ? " seg-on" : "")} onClick={() => setLine(p.id, { mode: "unite" })}>Par {p.unite.toLowerCase()}</button><button className={"seg-b" + (line?.mode === "paquet" ? " seg-on" : "")} onClick={() => setLine(p.id, { mode: "paquet" })} disabled={p.qteEmballage <= 1}>Emb. ×{p.qteEmballage}</button></div></>)}
            <label className="lbl">Acheté chez / facturer</label>
            <div className="ent-seg full">{ENT_ORDER.map((e) => <button key={e} className={"ent-b" + (curEnt === e ? " ent-on" : "")} style={curEnt === e ? { background: ENTITIES[e].color, borderColor: ENTITIES[e].color } : {}} onClick={() => hasVar ? setProdEntity(p.id, e) : setLine(p.id, { entity: e })}>{e}</button>)}</div>
            {!hasVar && cat?.num && (<><label className="lbl">N° matériel</label><input className="inp" placeholder="ex : #123" value={line?.assetNo || ""} onChange={(e) => setLine(p.id, { assetNo: e.target.value })} /></>)}
            <button className="btn btn-primary btn-block" onClick={() => setDetailsFor(null)}>OK</button>
          </div></div>);
      })()}
      {basketOpen && (
        <div className="modal-bg" onClick={() => setBasketOpen(false)}><div className="modal" onClick={(e) => e.stopPropagation()}>
          <div className="modal-h">Panier · {draft.lines.length} article(s)</div>
          <div className="basket-list">
            {draft.lines.length === 0 && <div className="empty sm">Panier vide.</div>}
            {[...draft.lines].sort((a, b) => byName(pById[a.productId] || { nom: "" }, pById[b.productId] || { nom: "" })).map((l) => { const p = pById[l.productId]; return (
              <div key={lineKey(l)} className="basket-row">
                <div className="basket-main"><div className="basket-name">{p?.nom}{l.variant ? " · " + l.variant : ""} <EntBadge e={l.entity} /></div><div className="basket-sub">Réf {p?.ref}</div></div>
                <Stepper value={l.qty} onChange={(v) => setLine(l.productId, { qty: v }, l.variant)} />
                <button className="del" onClick={() => removeKey(lineKey(l))}>✕</button>
              </div>); })}
          </div>
          {total > 0 && <div className="recap-total"><span>Total indicatif</span><b>{money(total)}</b></div>}
          <button className="btn btn-primary btn-block" onClick={() => setBasketOpen(false)}>Fermer</button>
        </div></div>)}
      {custOpen && (
        <div className="modal-bg" onClick={() => setCustOpen(false)}><div className="modal" onClick={(e) => e.stopPropagation()}>
          <div className="modal-h">Produit hors catalogue</div>
          <label className="lbl">Description</label><input className="inp" value={cf.nom} onChange={(e) => setCf({ ...cf, nom: e.target.value })} placeholder="Ex : Raccord Gardena 3/4" />
          <label className="lbl">Référence (facultatif)</label><input className="inp" value={cf.ref} onChange={(e) => setCf({ ...cf, ref: e.target.value })} />
          <div className="row3"><input className="inp" placeholder="Unité" value={cf.unite} onChange={(e) => setCf({ ...cf, unite: e.target.value })} /><input className="inp" placeholder="Prix" value={cf.prix} onChange={(e) => setCf({ ...cf, prix: e.target.value })} inputMode="decimal" /><input className="inp" placeholder="Qté" value={cf.qty} onChange={(e) => setCf({ ...cf, qty: e.target.value })} inputMode="numeric" /></div>
          <label className="lbl">Catégorie (position sur la fiche)</label><select className="inp" value={cf.catId} onChange={(e) => setCf({ ...cf, catId: e.target.value })}>{cats.map((c) => <option key={c.id} value={c.id}>{c.order}. {c.name}</option>)}</select>
          <label className="lbl">Acheté chez / facturer</label><div className="ent-seg full">{ENT_ORDER.map((e) => <button key={e} className={"ent-b" + (cf.entity === e ? " ent-on" : "")} style={cf.entity === e ? { background: ENTITIES[e].color, borderColor: ENTITIES[e].color } : {}} onClick={() => setCf({ ...cf, entity: e })}>{e}</button>)}</div>
          <button className="btn btn-primary btn-block" onClick={addCustom}>Ajouter à la demande</button>
          <button className="btn btn-ghost btn-block" onClick={() => setCustOpen(false)}>Annuler</button>
        </div></div>)}
    </div>
  );
}
function Stepper({ value, onChange }) {
  return (<div className="stepper"><button onClick={() => onChange(Math.max(0, value - 1))}>−</button><input value={value} onChange={(e) => { const n = parseInt(e.target.value.replace(/\D/g, ""), 10); onChange(isNaN(n) ? 0 : n); }} inputMode="numeric" /><button onClick={() => onChange(value + 1)}>+</button></div>);
}

/* ===================== Magasin (écran stock) ===================== */
function Magasin({ orders, onOpen, back }) {
  const [range, setRange] = useState("day");
  const now = new Date();
  function inRange(o) { const ref = o.departDate ? new Date(o.departDate + "T00:00") : new Date(o.createdAt); const diff = (ref - new Date(now.getFullYear(), now.getMonth(), now.getDate())) / 86400000; if (range === "day") return (diff >= 0 && diff < 1) || (ref < now); if (range === "week") return diff < 7; if (range === "month") return diff < 31; return true; }
  const list = orders.filter((o) => o.status !== "terminee" && inRange(o)).sort((a, b) => ((a.departDate || "") + (a.departTime || "")).localeCompare((b.departDate || "") + (b.departTime || "")));
  function exportXlsx() {
    const rows = [["Référence", "Chantier", "N° chantier", "Demandeur", "Date demande", "Date départ", "Heure départ", "Statut", "Magasinier", "Articles"]];
    list.forEach((o) => rows.push([o.ref, o.nom, o.numero, o.username, new Date(o.createdAt).toLocaleDateString("fr-BE"), fmtD(o.departDate), o.departTime || "", STATUS[o.status].label, o.assignedTo || "", o.lines.length]));
    const ws = XLSX.utils.aoa_to_sheet(rows); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Magasin");
    downloadBlob(new Blob([XLSX.write(wb, { type: "array", bookType: "xlsx" })], { type: "application/octet-stream" }), "pomproc_magasin.xlsx");
  }
  return (
    <div>
      <TopBar title="Magasin — à préparer" onBack={back} right={<button className="pill" onClick={exportXlsx}>Exporter</button>} />
      <div className="scroll">
        <div className="chips">{[["day", "Aujourd'hui"], ["week", "Semaine"], ["month", "Mois"], ["all", "Tout"]].map(([k, l]) => <button key={k} className={"chip" + (range === k ? " chip-on" : "")} onClick={() => setRange(k)}>{l}</button>)}</div>
        {list.length === 0 ? <div className="empty">Aucune demande à préparer sur cette période.</div> :
          <div className="mag-list">{list.map((o) => (
            <button key={o.id} className="mag-row" style={{ borderLeftColor: STATUS[o.status].color }} onClick={() => onOpen(o)}>
              <div className="mag-ref">{o.ref}</div>
              <div className="mag-mid">
                <div className="mag-name">{o.nom} <span className="ref-tag">N° {o.numero}</span></div>
                <div className="mag-sub">{o.username} · dem. {new Date(o.createdAt).toLocaleDateString("fr-BE")}{o.assignedTo ? " · " + o.assignedTo : ""}</div>
              </div>
              <div className="mag-dep"><span>Pour</span><b>{fmtD(o.departDate)}</b>{o.departTime ? <b className="hh">{o.departTime}</b> : null}</div>
              <StatusBadge st={o.status} />
            </button>))}</div>}
        <p className="pane-note">Vue destinée à l'écran du stock. Les demandes terminées passent dans « Terminées » (admin).</p>
      </div>
    </div>
  );
}
/* ===================== Préparation (magasinier) ===================== */
function Finished({ orders, onOpen, back }) {
  const [q, setQ] = useState("");
  const list = orders.filter((o) => o.status === "terminee" && (!q.trim() || (o.ref + " " + o.nom + " " + o.username + " " + (o.completedBy || "")).toLowerCase().includes(q.trim().toLowerCase()))).sort((a, b) => (b.completedAt || 0) - (a.completedAt || 0));
  return (
    <div><TopBar title="Demandes terminées" onBack={back} />
      <div className="scroll">
        <div className="search"><span>⌕</span><input placeholder="Réf, chantier, magasinier…" value={q} onChange={(e) => setQ(e.target.value)} />{q && <button onClick={() => setQ("")}>✕</button>}</div>
        {list.length === 0 ? <div className="empty">Aucune demande terminée.</div> :
          <div className="list">{list.map((o) => (
            <button key={o.id} className="row" onClick={() => onOpen(o)}>
              <div className="row-main"><div className="row-title">{o.nom} <span className="ref-tag">{o.ref}</span></div><div className="row-sub">{o.username} · préparée par <b>{o.completedBy || "—"}</b>{o.completedAt ? " · " + new Date(o.completedAt).toLocaleDateString("fr-BE") : ""}</div></div>
              <div className="row-right"><StatusBadge st={o.status} /><span className="chev">›</span></div>
            </button>))}</div>}
      </div>
    </div>
  );
}

function Prepare({ order, session, data, productById, catById, updateOrder, back, viewSheet, addItems }) {
  const [o, setO] = useState(order);
  const [magName, setMagName] = useState(order.completedBy || order.assignedTo || session.username);
  const photoRef = useRef();
  const [photoBusy, setPhotoBusy] = useState(false);
  const pById = useMemo(() => { const m = { ...productById }; (o.extras || []).forEach((e) => (m[e.id] = e)); return m; }, [productById, o.extras]);
  const cats = [...data.productCats].sort((a, b) => a.order - b.order);
  const grouped = cats.map((c) => ({ cat: c, lines: o.lines.filter((l) => pById[l.productId]?.catId === c.id).sort((a, b) => byName(pById[a.productId] || { nom: "" }, pById[b.productId] || { nom: "" })) })).filter((g) => g.lines.length);
  const stt = { ok: { l: "Préparé", c: "#128A52" }, missing: { l: "Manquant", c: "#D9453F" }, replaced: { l: "Remplacé", c: "#C9992B" } };
  const admin = session.role === "admin";
  const locked = o.status === "terminee" && !admin;
  function eff(l, pr) { const p = pById[l.productId]; if (!p) return 0; const factor = l.mode === "paquet" ? (p.qteEmballage || 1) : 1; const q = pr?.state === "missing" ? 0 : (pr?.qtyPrep != null ? pr.qtyPrep : l.qty); const price = (pr?.state === "replaced" && pr.replPrice != null && pr.replPrice !== "") ? Number(String(pr.replPrice).replace(",", ".")) || 0 : p.prix; return q * factor * price; }
  function recalc(ord) { let f = 0, d = 0; ord.lines.forEach((l) => { const pr = ord.prep[lineKey(l)]; const t = eff(l, pr); if (l.entity === "FFA") f += t; else if (l.entity === "DC") d += t; }); return { ...ord, billFFA: f, billDC: d }; }
  function commit(next) { if (locked) return; setO(next); updateOrder(next); }
  function persist() { if (locked) return; updateOrder(o); }
  function setPrepLocal(k, patch) { setO((x) => ({ ...x, prep: { ...x.prep, [k]: { ...(x.prep[k] || {}), ...patch } } })); }
  function setPrepCommit(k, patch) { commit(recalc({ ...o, prep: { ...o.prep, [k]: { ...(o.prep[k] || {}), ...patch } } })); }
  function setLineLocal(k, patch) { setO((x) => ({ ...x, lines: x.lines.map((l) => (lineKey(l) === k ? { ...l, ...patch } : l)) })); }
  function setLineCommit(k, patch) { commit(recalc({ ...o, lines: o.lines.map((l) => (lineKey(l) === k ? { ...l, ...patch } : l)) })); }
  async function take() { commit({ ...o, status: "en_cours", assignedTo: session.username }); }
  async function complete() { const n = recalc({ ...o, status: "terminee", completedBy: magName.trim() || session.username, completedAt: Date.now(), assignedTo: o.assignedTo || session.username }); setO(n); await updateOrder(n); back(); }
  async function saveEdits() { const n = recalc({ ...o, completedBy: magName.trim() || o.completedBy }); setO(n); await updateOrder(n); back(); }
  async function onPhoto(e) { if (locked) return; const file = e.target.files?.[0]; if (!file) return; try { setPhotoBusy(true); const url = await prepImageUpload(file); const next = { ...o, prepPhotos: [...(o.prepPhotos || []), url] }; setO(next); updateOrder(next); } catch (err) { window.alert("Photo impossible : " + (err?.message || err)); } finally { setPhotoBusy(false); e.target.value = ""; } }
  function removePhoto(i) { if (locked) return; const next = { ...o, prepPhotos: (o.prepPhotos || []).filter((_, idx) => idx !== i) }; setO(next); updateOrder(next); }
  const total = o.lines.reduce((s, l) => s + eff(l, o.prep[lineKey(l)]), 0);
  return (
    <div>
      <TopBar title={"Préparation " + o.ref} onBack={back} right={<button className="pill" onClick={viewSheet}>Fiche</button>} />
      <div className="scroll pb-bar">
        <div className="prep-head">
          <div className="prep-h-row"><b>{o.nom}</b> <span className="ref-tag">N° {o.numero}</span></div>
          <div className="prep-h-sub">Demandeur {o.username} · départ {fmtD(o.departDate)}{o.departTime ? " " + o.departTime : ""}</div>
          <button className="prep-add-btn" onClick={addItems}>＋ Ajouter un article</button>
          <StatusBadge st={o.status} />
        </div>
        {o.status === "envoyee" && <button className="btn btn-primary btn-block" onClick={take}>Prendre en charge</button>}
        {o.status === "terminee" && <div className={"lock-banner" + (locked ? "" : " editable")}>✔ Terminée par {o.completedBy || "—"} — {locked ? "verrouillée (seul un administrateur peut modifier)" : "modification administrateur activée"}</div>}
        {grouped.map((g) => (
          <div key={g.cat.id} className="prep-cat">
            <div className="prep-cat-h" style={{ background: g.cat.color }}><CatIcon icon={g.cat.icon} color="#fff" size={15} /> {g.cat.order}. {g.cat.name}</div>
            {g.lines.map((l) => { const p = pById[l.productId]; const lk = lineKey(l); const pr = o.prep[lk] || {}; const delivered = pr.qtyPrep != null ? pr.qtyPrep : l.qty; return (
              <div key={lk} className={"prep-line" + (pr.state ? " treated" : "")}>
                <div className="prep-line-top"><div className="prep-info"><div className="prep-name">{p.nom}{l.variant ? " · " + l.variant : ""} <EntBadge e={l.entity} />{l.addedAt ? <span className="add-tag">＋{new Date(l.addedAt).toLocaleDateString("fr-BE")}</span> : null}</div><div className="prep-sub">Demandé {l.qty} × {l.mode === "paquet" ? "emb. " + p.qteEmballage : p.unite.toLowerCase()} · Réf {p.ref}{p.prix ? " · " + money(p.prix) : ""}</div></div></div>
                <div className="prep-states">{["ok", "missing", "replaced"].map((s) => <button key={s} className={"pstate" + (pr.state === s ? " on" : "")} style={pr.state === s ? { background: stt[s].c, borderColor: stt[s].c, color: "#fff" } : { color: stt[s].c }} disabled={locked} onClick={() => setPrepCommit(lk, { state: s, qtyPrep: s === "missing" ? 0 : (pr.qtyPrep != null ? pr.qtyPrep : l.qty) })}>{stt[s].l}</button>)}</div>
                {pr.state && pr.state !== "missing" && <div className="prep-qty"><span>Livré</span><Stepper value={delivered} onChange={(v) => setPrepCommit(lk, { qtyPrep: v })} /><span className="qty-of">/ {l.qty} demandé{p.prix ? " · " + money(delivered * (l.mode === "paquet" ? (p.qteEmballage || 1) : 1) * ((pr.state === "replaced" && pr.replPrice) ? (Number(String(pr.replPrice).replace(",", ".")) || 0) : p.prix)) : ""}</span></div>}
                {pr.state === "replaced" && <div className="repl-box">
                  <div className="repl-search"><span>⌕</span><input placeholder="Remplacé par… (chercher ou saisir)" disabled={locked} value={pr.repl || ""} onChange={(e) => setPrepLocal(lk, { repl: e.target.value, replChosen: false })} onBlur={persist} /></div>
                  {!locked && pr.repl && !pr.replChosen && (() => { const s = pr.repl.trim().toLowerCase(); const res = data.products.filter((pp) => pp.nom.toLowerCase().includes(s) || String(pp.ref).toLowerCase().includes(s)).slice(0, 5); return res.length ? <div className="repl-res">{res.map((pp) => <button key={pp.id} onClick={() => setPrepCommit(lk, { repl: pp.nom, replRef: pp.ref, replPrice: pp.prix, replChosen: true })}>{pp.nom} <small>{pp.ref} · {money(pp.prix)}</small></button>)}</div> : null; })()}
                  <input className="asset-inp price-inp" placeholder="Prix unit. (€)" disabled={locked} value={pr.replPrice ?? ""} inputMode="decimal" onChange={(e) => setPrepLocal(lk, { replPrice: e.target.value })} onBlur={() => commit(recalc(o))} />
                </div>}
                <div className="prep-edit">
                  {g.cat.num && <div className="pe-field"><span>N° matériel</span><input className="asset-inp mini" placeholder="ex : #123" disabled={locked} value={l.assetNo || ""} onChange={(e) => setLineLocal(lk, { assetNo: e.target.value })} onBlur={persist} /></div>}
                  <div className="pe-field"><span>Facturer</span><div className="ent-seg">{ENT_ORDER.map((e) => <button key={e} className={"ent-b" + (l.entity === e ? " ent-on" : "")} style={l.entity === e ? { background: ENTITIES[e].color, borderColor: ENTITIES[e].color } : {}} disabled={locked} onClick={() => setLineCommit(lk, { entity: e })}>{e}</button>)}</div></div>
                </div>
              </div>); })}
          </div>
        ))}
        <div className="prep-total">Total (livré) <b>{money(total)}</b></div>
        <label className="lbl">Remarque du magasinier</label>
        <textarea className="inp ta" disabled={locked} value={o.magasinierNote || ""} onChange={(e) => setO({ ...o, magasinierNote: e.target.value })} onBlur={persist} placeholder="Commentaire, ruptures, remplacements globaux…" />
        <label className="lbl">Nom du magasinier</label>
        <input className="inp" disabled={locked} value={magName} onChange={(e) => setMagName(e.target.value)} />
        <label className="lbl">Photos de la préparation (état du matériel)</label>
        <div className="photo-grid">
          {(o.prepPhotos || []).map((src, i) => <div key={i} className="photo-thumb"><img src={src} alt="" />{!locked && <button onClick={() => removePhoto(i)}>✕</button>}</div>)}
          {!locked && <label className="photo-add">{photoBusy ? "…" : "+ Photo"}<input ref={photoRef} type="file" accept="image/*" capture="environment" style={{ display: "none" }} onChange={onPhoto} /></label>}
        </div>
      </div>
      <div className="bottombar no-print"><div className="bb-total" />{locked ? <div className="lock-note">🔒 Préparation verrouillée</div> : <button className="btn btn-primary" onClick={o.status === "terminee" ? saveEdits : complete}>{o.status === "terminee" ? "Enregistrer les modifications" : "Terminer la préparation"}</button>}</div>
    </div>
  );
}
/* ============================ Templates ============================ */
function Templates({ session, data, saveData, productById, catById, back }) {
  const fileRef = useRef();
  const tpls = data.templates || [];
  const [editing, setEditing] = useState(null); // template object or 'new'
  const canEdit = (t) => t.ownerUsername === session.username || session.role === "admin";
  function del(t) { if (!canEdit(t)) return; if (!window.confirm("Supprimer « " + t.name + " » ?")) return; saveData({ ...data, templates: tpls.filter((x) => x !== t) }); }
  function exportXlsx(list) {
    const rows = [["Template", "Réf", "Description", "Qté", "Mode"]];
    list.forEach((t) => t.lines.forEach((l) => { const p = productById[l.productId]; if (p) rows.push([t.name, p.ref, p.nom, l.qty, l.mode]); }));
    const ws = XLSX.utils.aoa_to_sheet(rows); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Templates");
    downloadBlob(new Blob([XLSX.write(wb, { type: "array", bookType: "xlsx" })], { type: "application/octet-stream" }), "pomproc_templates.xlsx");
  }
  async function onImport(e) {
    const file = e.target.files?.[0]; if (!file) return;
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
      const refMap = {}; data.products.forEach((p) => (refMap[String(p.ref).trim()] = p.id));
      const groups = {}; let skipped = 0;
      rows.slice(1).forEach((r) => { if (!r || !r[0]) return; const [name, ref, , qty, mode] = r; const pid = refMap[String(ref).trim()]; if (!pid) { skipped++; return; } (groups[name] = groups[name] || []).push({ productId: pid, qty: Number(qty) || 1, mode: mode === "paquet" ? "paquet" : "unite" }); });
      const imported = Object.entries(groups).map(([name, lines]) => ({ id: uid("tpl_"), name, ownerUsername: session.username, shared: true, lines, createdAt: Date.now() }));
      if (!imported.length) return window.alert("Aucun template valide trouvé.");
      await saveData({ ...data, templates: [...tpls, ...imported] });
      window.alert(imported.length + " template(s) importé(s)." + (skipped ? " " + skipped + " ligne(s) ignorée(s)." : ""));
    } catch { window.alert("Fichier illisible."); }
    e.target.value = "";
  }
  if (editing) return <TemplateEditor tpl={editing === "new" ? null : editing} data={data} saveData={saveData} productById={productById} catById={catById} session={session} back={() => setEditing(null)} />;
  return (
    <div>
      <TopBar title="Templates" onBack={back} right={<button className="pill" onClick={() => fileRef.current?.click()}>Importer</button>} />
      <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={onImport} />
      <div className="scroll">
        <button className="btn btn-primary btn-block" onClick={() => setEditing("new")}>+ Créer un template</button>
        {tpls.length > 0 && <button className="btn btn-ghost btn-block" onClick={() => exportXlsx(tpls)}>⭳ Exporter tout (Excel)</button>}
        {tpls.length === 0 ? <div className="empty">Aucun template.</div> :
          <div className="list">{tpls.map((t) => (
            <div key={t.id} className="row static">
              <div className="row-main" onClick={() => canEdit(t) && setEditing(t)} style={{ cursor: canEdit(t) ? "pointer" : "default" }}><div className="row-title">{t.name}</div><div className="row-sub">{t.lines.length} article(s) · {t.ownerUsername}</div></div>
              <button className="mini-btn" onClick={() => exportXlsx([t])}>⭳</button>
              {canEdit(t) && <button className="mini-btn" onClick={() => setEditing(t)}>✎</button>}
              {canEdit(t) && <button className="del big" onClick={() => del(t)}>✕</button>}
            </div>))}</div>}
      </div>
    </div>
  );
}

function TemplateEditor({ tpl, data, saveData, productById, catById, session, back }) {
  const cats = useMemo(() => [...data.productCats].sort((a, b) => a.order - b.order), [data.productCats]);
  const [name, setName] = useState(tpl?.name || "");
  const [lines, setLines] = useState(tpl ? tpl.lines.map((l) => ({ ...l })) : []);
  const [openCat, setOpenCat] = useState(cats[0]?.id || "");
  const [q, setQ] = useState("");
  const lineFor = (pid) => lines.find((l) => l.productId === pid);
  function setLine(pid, qty) { setLines((ls) => { const i = ls.findIndex((l) => l.productId === pid); let n; if (i === -1) n = [...ls, { productId: pid, qty, mode: "unite" }]; else { n = ls.slice(); n[i] = { ...n[i], qty }; } return n.filter((l) => l.qty > 0); }); }
  const shown = useMemo(() => { let list = data.products; if (q.trim()) { const s = q.trim().toLowerCase(); list = list.filter((p) => p.nom.toLowerCase().includes(s) || String(p.ref).toLowerCase().includes(s)); } else list = list.filter((p) => p.catId === openCat); return [...list].sort(byName); }, [q, openCat, data.products]);
  function save() {
    if (!name.trim() || !lines.length) return window.alert("Nom et au moins un produit requis.");
    if (tpl) saveData({ ...data, templates: data.templates.map((t) => t.id === tpl.id ? { ...t, name: name.trim(), lines } : t) });
    else saveData({ ...data, templates: [...(data.templates || []), { id: uid("tpl_"), name: name.trim(), ownerUsername: session.username, shared: true, lines, createdAt: Date.now() }] });
    back();
  }
  return (
    <div>
      <TopBar title={tpl ? "Modifier template" : "Nouveau template"} onBack={back} right={<button className="pill" onClick={save}>Enregistrer</button>} />
      <div className="scroll pb-bar">
        <label className="lbl">Nom du template</label><input className="inp" value={name} onChange={(e) => setName(e.target.value)} placeholder="Ex : Nettoyage bureaux" />
        <div className="search" style={{ marginTop: 12 }}><span>⌕</span><input placeholder="Rechercher…" value={q} onChange={(e) => setQ(e.target.value)} />{q && <button onClick={() => setQ("")}>✕</button>}</div>
        {!q && <div className="catwrap">{cats.map((c) => <button key={c.id} className="ccat" style={openCat === c.id ? { background: c.color, borderColor: c.color, color: "#fff" } : { color: c.color, borderColor: c.color }} onClick={() => setOpenCat(c.id)}><CatIcon icon={c.icon} color={openCat === c.id ? "#fff" : c.color} size={15} /><span>{c.order}. {c.name}</span></button>)}</div>}
        <div className="prodlist">{shown.map((p) => { const l = lineFor(p.id); const cat = catById[p.catId]; return (
          <div key={p.id} className={"prod" + (l ? " prod-on" : "")} style={l ? { borderColor: cat?.color } : {}}>
            <div className="prod-top"><div className="prod-thumb" style={{ borderColor: cat?.color }}>{p.image ? <img src={p.image} alt="" /> : <CatIcon icon={cat?.icon} color={cat?.color} size={20} />}</div>
              <div className="prod-info"><div className="prod-name">{p.nom}</div><div className="prod-meta">Réf {p.ref} · {p.unite}</div></div>
              <Stepper value={l ? l.qty : 0} onChange={(v) => setLine(p.id, v)} /></div>
          </div>); })}</div>
      </div>
      <div className="bottombar no-print"><div className="bb-total"><small>{lines.length} article(s)</small></div><button className="btn btn-primary" onClick={save}>Enregistrer</button></div>
    </div>
  );
}

/* ===================== Admin : listes ===================== */
function AllOrders({ orders, onOpen, back }) {
  const [q, setQ] = useState("");
  const list = orders.filter((o) => !q.trim() || (o.ref + " " + o.nom + " " + o.numero + " " + o.username).toLowerCase().includes(q.trim().toLowerCase()));
  return (
    <div><TopBar title="Toutes les demandes" onBack={back} />
      <div className="scroll">
        <div className="search"><span>⌕</span><input placeholder="Réf, chantier, employé…" value={q} onChange={(e) => setQ(e.target.value)} />{q && <button onClick={() => setQ("")}>✕</button>}</div>
        {list.length === 0 ? <div className="empty">Aucune demande.</div> : <div className="list">{list.map((o) => <OrderRow key={o.id} o={o} showUser onClick={() => onOpen(o)} />)}</div>}
      </div>
    </div>
  );
}
function Billing({ orders, productById, session, updateOrder, onOpen, back }) {
  const isAdmin = session?.role === "admin";
  const [mode, setMode] = useState("month");
  const [mf, setMf] = useState("");
  const monthsAvail = useMemo(() => Array.from(new Set(orders.map((o) => monthKey(o.createdAt)))).sort((a, b) => b.localeCompare(a)), [orders]);
  function stats(o) {
    const f0 = o.billFFA0 ?? o.billFFA ?? 0, d0 = o.billDC0 ?? o.billDC ?? 0, f = o.billFFA ?? 0, d = o.billDC ?? 0;
    const iss = Object.values(o.prep || {}).filter((v) => v.state && v.state !== "ok");
    const repl = iss.filter((v) => v.state === "replaced").length, miss = iss.filter((v) => v.state === "missing").length;
    const dF = f - f0, dD = d - d0;
    return { f, d, dF, dD, repl, miss, changed: repl > 0 || miss > 0 || dF !== 0 || dD !== 0 };
  }
  const scoped = useMemo(() => orders.filter((o) => !mf || monthKey(o.createdAt) === mf), [orders, mf]);
  const bt = (o) => (o.billFFA || 0) + (o.billDC || 0);
  const isPaid = (o) => o.paymentStatus === "paye";
  const grand = useMemo(() => { let FFA = 0, DC = 0, paid = 0, unpaid = 0; scoped.forEach((o) => { FFA += o.billFFA || 0; DC += o.billDC || 0; const t = bt(o); if (isPaid(o)) paid += t; else unpaid += t; }); return { FFA, DC, total: FFA + DC, paid, unpaid }; }, [scoped]);
  const keyOfMode = (o, m) => m === "month" ? monthKey(o.createdAt) : m === "demandeur" ? (o.username || "—") : m === "chef" ? (o.responsable || "—") : (o.status === "terminee" ? (o.completedBy || "—") : null);
  const groups = useMemo(() => {
    const g = {};
    scoped.forEach((o) => { const key = keyOfMode(o, mode); if (key == null) return; const s = stats(o); const e = (g[key] = g[key] || { FFA: 0, DC: 0, count: 0, changed: 0, dF: 0, dD: 0, paid: 0, unpaid: 0 }); e.FFA += s.f; e.DC += s.d; e.count++; if (s.changed) e.changed++; e.dF += s.dF; e.dD += s.dD; if (isPaid(o)) e.paid += bt(o); else e.unpaid += bt(o); });
    const arr = Object.entries(g);
    if (mode === "month") arr.sort((a, b) => b[0].localeCompare(a[0])); else arr.sort((a, b) => (b[1].FFA + b[1].DC) - (a[1].FFA + a[1].DC));
    return arr;
  }, [scoped, mode]);
  const billable = useMemo(() => scoped.filter((o) => bt(o) > 0).sort((a, b) => (isPaid(a) === isPaid(b) ? b.createdAt - a.createdAt : (isPaid(a) ? 1 : -1))), [scoped]);
  const changedList = useMemo(() => scoped.map((o) => ({ o, s: stats(o) })).filter((x) => x.s.changed).sort((a, b) => b.o.createdAt - a.o.createdAt), [scoped]);
  const label = (k) => (mode === "month" ? monthLabel(k) : k);
  const signed = (n) => (n > 0 ? "+" : "") + money(n);
  function togglePay(o) { if (!isAdmin) return; updateOrder({ ...o, paymentStatus: isPaid(o) ? "non_paye" : "paye" }); }
  function exportXlsx() {
    const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
    const wb = XLSX.utils.book_new();
    const aggSheet = (keyOf, keyLabel, isMonth, terminOnly) => {
      const g = {};
      scoped.forEach((o) => { if (terminOnly && o.status !== "terminee") return; const k = keyOf(o); if (k == null) return; const s = stats(o); const e = (g[k] = g[k] || { FFA: 0, DC: 0, count: 0, changed: 0, dF: 0, dD: 0, paid: 0, unpaid: 0 }); e.FFA += s.f; e.DC += s.d; e.count++; if (s.changed) e.changed++; e.dF += s.dF; e.dD += s.dD; if (isPaid(o)) e.paid += bt(o); else e.unpaid += bt(o); });
      const entries = Object.entries(g);
      if (isMonth) entries.sort((a, b) => b[0].localeCompare(a[0])); else entries.sort((a, b) => (b[1].FFA + b[1].DC) - (a[1].FFA + a[1].DC));
      const rows = [[keyLabel, "FFA (€)", "DC (€)", "Total (€)", "Payé (€)", "Non payé (€)", "Demandes", "Modifiées", "Δ FFA (€)", "Δ DC (€)"]];
      const t = { FFA: 0, DC: 0, count: 0, changed: 0, dF: 0, dD: 0, paid: 0, unpaid: 0 };
      entries.forEach(([k, v]) => { rows.push([isMonth ? monthLabel(k) : k, r2(v.FFA), r2(v.DC), r2(v.FFA + v.DC), r2(v.paid), r2(v.unpaid), v.count, v.changed, r2(v.dF), r2(v.dD)]); Object.keys(t).forEach((kk) => (t[kk] += v[kk])); });
      rows.push(["TOTAL", r2(t.FFA), r2(t.DC), r2(t.FFA + t.DC), r2(t.paid), r2(t.unpaid), t.count, t.changed, r2(t.dF), r2(t.dD)]);
      return XLSX.utils.aoa_to_sheet(rows);
    };
    const sumSheet = XLSX.utils.aoa_to_sheet([["Récapitulatif" + (mf ? " — " + monthLabel(mf) : " — tous les mois")], [], ["Total facturé (€)", r2(grand.total)], ["dont FFA (€)", r2(grand.FFA)], ["dont DC (€)", r2(grand.DC)], [], ["Payé (€)", r2(grand.paid)], ["Non payé (€)", r2(grand.unpaid)]]);
    XLSX.utils.book_append_sheet(wb, sumSheet, "Récapitulatif");
    XLSX.utils.book_append_sheet(wb, aggSheet((o) => monthKey(o.createdAt), "Mois", true, false), "Par mois");
    XLSX.utils.book_append_sheet(wb, aggSheet((o) => o.username || "—", "Demandeur", false, false), "Par demandeur");
    XLSX.utils.book_append_sheet(wb, aggSheet((o) => o.responsable || "—", "Chef d'équipe", false, false), "Par chef d'équipe");
    XLSX.utils.book_append_sheet(wb, aggSheet((o) => o.completedBy || "—", "Magasinier", false, true), "Par magasinier");
    const det = [["Référence", "Date", "Mois", "Chantier", "N° chantier", "Demandeur", "Chef d'équipe", "Magasinier", "Statut", "Paiement", "FFA (€)", "DC (€)", "FFA origine", "DC origine", "Δ FFA", "Δ DC", "Remplacements", "Manquants", "Total (€)"]];
    scoped.slice().sort((a, b) => b.createdAt - a.createdAt).forEach((o) => { const s = stats(o); det.push([o.ref, new Date(o.createdAt).toLocaleDateString("fr-BE"), monthKey(o.createdAt), o.nom, o.numero, o.username, o.responsable || "", o.completedBy || "", STATUS[o.status]?.label || o.status, isPaid(o) ? "Payé" : "Non payé", r2(s.f), r2(s.d), r2(o.billFFA0 ?? o.billFFA ?? 0), r2(o.billDC0 ?? o.billDC ?? 0), r2(s.dF), r2(s.dD), s.repl, s.miss, r2(o.total)]); });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(det), "Détail demandes");
    const mods = [["Référence", "Demandeur", "Magasinier", "Produit", "Réf produit", "Changement", "Remplacé par", "Entité facturée", "Montant ligne (€)"]];
    scoped.forEach((o) => { const lbk = {}; o.lines.forEach((l) => (lbk[lineKey(l)] = l)); const ex = {}; (o.extras || []).forEach((e) => (ex[e.id] = e)); Object.entries(o.prep || {}).forEach(([key, v]) => { if (!v.state || v.state === "ok") return; const line = lbk[key]; const p = productById[line?.productId] || ex[line?.productId]; const amt = line ? lineTotal(line, p) : 0; mods.push([o.ref, o.username, o.completedBy || "", (p?.nom || key) + (line?.variant ? " " + line.variant : ""), p?.ref || "", v.state === "missing" ? "Manquant" : "Remplacé", v.repl || "", line?.entity || "", r2(amt)]); }); });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mods), "Modifications");
    downloadBlob(new Blob([XLSX.write(wb, { type: "array", bookType: "xlsx" })], { type: "application/octet-stream" }), "pomproc_facturation" + (mf ? "_" + mf : "") + ".xlsx");
  }
  return (
    <div><TopBar title="Facturation" onBack={back} right={<button className="pill" onClick={exportXlsx}>Exporter</button>} />
      <div className="scroll">
        <select className="inp" value={mf} onChange={(e) => setMf(e.target.value)} style={{ marginBottom: 12 }}><option value="">Tous les mois</option>{monthsAvail.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}</select>
        <div className="grand-card">
          <div className="grand-h">Total facturé{mf ? " · " + monthLabel(mf) : ""}</div>
          <div className="grand-total">{money(grand.total)}</div>
          <div className="grand-split"><span className="ebadge" style={{ background: ENTITIES.FFA.color }}>FFA</span> {money(grand.FFA)} &nbsp;·&nbsp; <span className="ebadge" style={{ background: ENTITIES.DC.color }}>DC</span> {money(grand.DC)}</div>
          <div className="grand-pay"><div className="gp paid"><span>Payé</span><b>{money(grand.paid)}</b></div><div className="gp unpaid"><span>Non payé</span><b>{money(grand.unpaid)}</b></div></div>
        </div>
        <div className="chips">{[["month", "Par mois"], ["demandeur", "Par demandeur"], ["chef", "Par chef d'équipe"], ["magasinier", "Par magasinier"]].map(([k, l]) => <button key={k} className={"chip" + (mode === k ? " chip-on" : "")} onClick={() => setMode(k)}>{l}</button>)}</div>
        <p className="pane-note">Facturé à FFA et DC{mode === "magasinier" ? " (préparations terminées)" : ""}. IIS (interne) non facturé. « Modifs magasinier » = écart de facturation dû aux changements en préparation.</p>
        {groups.length === 0 ? <div className="empty">Aucune donnée.</div> :
          groups.map(([k, v]) => (
            <div key={k} className="bill-card">
              <div className="bill-month">{label(k)}<small>{v.count} demande(s){v.changed ? " · " + v.changed + " modifiée(s)" : ""}</small></div>
              <div className="bill-rows">
                <div className="bill-r"><span className="ebadge" style={{ background: ENTITIES.FFA.color }}>FFA</span><b>{money(v.FFA)}</b></div>
                <div className="bill-r"><span className="ebadge" style={{ background: ENTITIES.DC.color }}>DC</span><b>{money(v.DC)}</b></div>
                <div className="bill-r total"><span>Total facturé</span><b>{money(v.FFA + v.DC)}</b></div>
                {v.paid ? <div className="bill-r paid-r"><span>Payé</span><b>{money(v.paid)}</b></div> : null}
                {v.unpaid ? <div className="bill-r unpaid-r"><span>Non payé</span><b>{money(v.unpaid)}</b></div> : null}
                {(v.dF || v.dD) ? <div className="bill-r delta"><span>dont modifs magasinier</span><b>{v.dF ? "FFA " + signed(v.dF) : ""}{v.dF && v.dD ? " · " : ""}{v.dD ? "DC " + signed(v.dD) : ""}</b></div> : null}
              </div>
            </div>))}
        <div className="sec-h">Paiements {isAdmin ? "" : "(lecture seule)"}</div>
        {billable.length === 0 ? <div className="empty sm">Aucune demande facturable.</div> :
          <div className="list">{billable.map((o) => (
            <div key={o.id} className="pay-row">
              <button className="pay-main" onClick={() => onOpen(o)}>
                <div className="row-title">{o.nom} <span className="ref-tag">{o.ref}</span></div>
                <div className="row-sub">{o.username}{o.responsable ? " · " + o.responsable : ""} · {money(bt(o))}</div>
              </button>
              {isAdmin
                ? <button className={"pay-toggle" + (isPaid(o) ? " on" : "")} onClick={() => togglePay(o)}>{isPaid(o) ? "✓ Payé" : "Non payé"}</button>
                : <span className={"pay-badge" + (isPaid(o) ? " on" : "")}>{isPaid(o) ? "Payé" : "Non payé"}</span>}
            </div>))}</div>}
        {changedList.length > 0 && (
          <div>
            <div className="sec-h">Demandes avec modifications</div>
            <div className="list">{changedList.map(({ o, s }) => (
              <button key={o.id} className="row" onClick={() => onOpen(o)}>
                <div className="row-main">
                  <div className="row-title">{o.nom} <span className="ref-tag">{o.ref}</span></div>
                  <div className="row-sub">{o.username}{o.completedBy ? " → " + o.completedBy : ""}{s.repl ? " · " + s.repl + " rempl." : ""}{s.miss ? " · " + s.miss + " manq." : ""}{(s.dF || s.dD) ? " · Δ " + money(s.dF + s.dD) : ""}</div>
                </div>
                <div className="row-right"><b className="bill-tot">{money(s.f + s.d)}</b><span className="chev">›</span></div>
              </button>))}</div>
          </div>)}
      </div>
    </div>
  );
}

/* ===================== Admin : analyse produits ===================== */
function Donut({ segments, size = 148, thick = 24 }) {
  const r = (size - thick) / 2, C = 2 * Math.PI * r, cx = size / 2;
  const total = segments.reduce((s, x) => s + x.val, 0) || 1;
  let acc = 0;
  return (
    <svg width={size} height={size} viewBox={"0 0 " + size + " " + size} style={{ flex: "none" }}>
      <g transform={"rotate(-90 " + cx + " " + cx + ")"}>
        <circle cx={cx} cy={cx} r={r} fill="none" stroke="#e6f0ea" strokeWidth={thick} />
        {segments.map((s, i) => { const len = s.val / total * C; const el = <circle key={i} cx={cx} cy={cx} r={r} fill="none" stroke={s.color} strokeWidth={thick} strokeDasharray={len + " " + (C - len)} strokeDashoffset={-acc} />; acc += len; return el; })}
      </g>
      <text x={cx} y={cx - 2} textAnchor="middle" fontSize="22" fontWeight="800" fill="#06301f">{total}</text>
      <text x={cx} y={cx + 15} textAnchor="middle" fontSize="9" fill="#5f7a6b">articles</text>
    </svg>
  );
}

function Analytics({ orders, data, productById, catById, back }) {
  const [range, setRange] = useState("all");
  const [sortBy, setSortBy] = useState("freq");
  const filtered = useMemo(() => { const now = Date.now(); return orders.filter((o) => range === "all" ? true : (now - o.createdAt) <= (range === "90" ? 90 : 30) * 86400000); }, [orders, range]);
  const A = useMemo(() => {
    const freq = {}, qty = {}, pairs = {}, replaced = {}, missing = {}, catAgg = {};
    const entTot = { FFA: { qty: 0, cost: 0, lines: 0 }, IIS: { qty: 0, cost: 0, lines: 0 }, DC: { qty: 0, cost: 0, lines: 0 } };
    const crossItems = {}; const entFlow = {};
    filtered.forEach((o) => {
      o.lines.forEach((l) => { freq[l.productId] = (freq[l.productId] || 0) + 1; const p = productById[l.productId]; const eff = l.qty * (l.mode === "paquet" ? (p?.qteEmballage || 1) : 1); qty[l.productId] = (qty[l.productId] || 0) + eff; const cid = p?.catId; if (cid) catAgg[cid] = (catAgg[cid] || 0) + 1; const billed = l.entity || p?.entity || "IIS"; const owner = p?.entity || "IIS"; const cost = lineTotal(l, p); (entTot[billed] = entTot[billed] || { qty: 0, cost: 0, lines: 0 }); entTot[billed].qty += eff; entTot[billed].cost += cost; entTot[billed].lines++; const fk = owner + ">" + billed; (entFlow[fk] = entFlow[fk] || { owner, billed, qty: 0, cost: 0 }); entFlow[fk].qty += eff; entFlow[fk].cost += cost; if (billed !== "IIS") { const ci = (crossItems[l.productId] = crossItems[l.productId] || { pid: l.productId, entity: billed, qty: 0, cost: 0 }); ci.qty += eff; ci.cost += cost; ci.entity = billed; } });
      const uniq = [...new Set(o.lines.map((l) => l.productId))];
      for (let i = 0; i < uniq.length; i++) for (let j = i + 1; j < uniq.length; j++) { const a = uniq[i], b = uniq[j]; const k = a < b ? a + "|" + b : b + "|" + a; pairs[k] = (pairs[k] || 0) + 1; }
      Object.entries(o.prep || {}).forEach(([key, v]) => { const pid = key.split("::")[0]; if (v.state === "replaced") { const r = (replaced[pid] = replaced[pid] || { count: 0, repls: {} }); r.count++; const t = (v.repl || "").trim(); if (t) r.repls[t] = (r.repls[t] || 0) + 1; } else if (v.state === "missing") missing[pid] = (missing[pid] || 0) + 1; });
    });
    const top = Object.entries(freq).map(([pid, f]) => ({ pid, f, q: qty[pid] || 0 }));
    const pairsArr = Object.entries(pairs).map(([k, c]) => { const [a, b] = k.split("|"); return { a, b, c }; }).sort((x, y) => y.c - x.c).slice(0, 8);
    const replacedArr = Object.entries(replaced).map(([pid, r]) => ({ pid, count: r.count, top: Object.entries(r.repls).sort((x, y) => y[1] - x[1])[0]?.[0] || "" })).sort((x, y) => y.count - x.count).slice(0, 8);
    const missingArr = Object.entries(missing).map(([pid, c]) => ({ pid, c })).sort((x, y) => y.c - x.c).slice(0, 8);
    const never = data.products.filter((p) => !freq[p.id]);
    const catShare = data.productCats.map((c) => ({ catId: c.id, name: c.name, color: c.color, val: catAgg[c.id] || 0 })).filter((x) => x.val > 0).sort((a, b) => b.val - a.val);
    const heatIds = Object.entries(freq).sort((a, b) => b[1] - a[1]).slice(0, 7).map((e) => e[0]);
    const matrix = heatIds.map((a, i) => heatIds.map((b, j) => { if (i === j) return 0; const k = a < b ? a + "|" + b : b + "|" + a; return pairs[k] || 0; }));
    const heatMax = Math.max(1, ...matrix.flat());
    const crossArr = Object.values(crossItems).sort((a, b) => b.cost - a.cost || b.qty - a.qty).slice(0, 10);
    const flowArr = Object.values(entFlow).filter((x) => x.qty > 0).sort((a, b) => b.cost - a.cost);
    return { top, pairsArr, replacedArr, missingArr, never, count: filtered.length, catShare, heat: { ids: heatIds, matrix, max: heatMax }, entTot, crossArr, flowArr };
  }, [filtered, productById, data.products, data.productCats]);

  const extraById = useMemo(() => { const m = {}; orders.forEach((o) => (o.extras || []).forEach((e) => (m[e.id] = e))); return m; }, [orders]);
  const nm = (pid) => (productById[pid] || extraById[pid])?.nom || pid;
  const col = (pid) => catById[(productById[pid] || extraById[pid])?.catId]?.color || "#999";
  const topSorted = [...A.top].sort((a, b) => (sortBy === "freq" ? b.f - a.f : b.q - a.q)).slice(0, 12);
  const maxVal = topSorted.length ? (sortBy === "freq" ? topSorted[0].f : topSorted[0].q) : 1;
  const totalVal = A.top.reduce((s, t) => s + (sortBy === "freq" ? t.f : t.q), 0) || 1;
  let run = 0; const rows = topSorted.map((t) => { const val = sortBy === "freq" ? t.f : t.q; run += val; return { pid: t.pid, val, cum: Math.round(run / totalVal * 100) }; });
  const catTot = A.catShare.reduce((s, x) => s + x.val, 0) || 1;

  function exportXlsx() {
    const wb = XLSX.utils.book_new();
    const s1 = [["Rang", "Produit", "Réf", "Catégorie", "Demandes", "Quantité"]];
    [...A.top].sort((a, b) => b.f - a.f).forEach((t, i) => { const p = productById[t.pid]; s1.push([i + 1, p?.nom || t.pid, p?.ref || "", catById[p?.catId]?.name || "", t.f, t.q]); });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s1), "Top produits");
    const sc = [["Catégorie", "Articles"]]; A.catShare.forEach((c) => sc.push([c.name, c.val]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sc), "Par catégorie");
    const s2 = [["Produit A", "Produit B", "Demandes communes"]]; A.pairsArr.forEach((pr) => s2.push([nm(pr.a), nm(pr.b), pr.c]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s2), "Combinaisons");
    const s3 = [["Produit", "Réf", "Remplacements", "Souvent remplacé par"]]; A.replacedArr.forEach((r) => { const p = productById[r.pid]; s3.push([p?.nom || r.pid, p?.ref || "", r.count, r.top]); });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s3), "Remplacements");
    const s4 = [["Produit", "Réf", "Manquants"]]; A.missingArr.forEach((m) => { const p = productById[m.pid]; s4.push([p?.nom || m.pid, p?.ref || "", m.c]); });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s4), "Manquants");
    const s5 = [["Produit jamais commandé", "Réf", "Catégorie"]]; A.never.forEach((p) => s5.push([p.nom, p.ref, catById[p.catId]?.name || ""]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s5), "Jamais commandés");
    const se1 = [["Entité (source / facturée)", "Lignes", "Quantité", "Montant (€)"]]; ENT_ORDER.forEach((e) => { const v = A.entTot[e] || { qty: 0, cost: 0, lines: 0 }; se1.push([e, v.lines, v.qty, Math.round(v.cost * 100) / 100]); });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(se1), "Flux par entité");
    const se2 = [["Propriétaire (owner)", "Facturé à", "Quantité", "Montant (€)"]]; A.flowArr.forEach((f) => se2.push([f.owner, f.billed, f.qty, Math.round(f.cost * 100) / 100]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(se2), "Matrice de flux");
    const se3 = [["Produit", "Réf", "Entité source", "Quantité", "Montant (€)"]]; A.crossArr.forEach((c) => { const p = productById[c.pid] || extraById[c.pid]; se3.push([p?.nom || c.pid, p?.ref || "", c.entity, c.qty, Math.round(c.cost * 100) / 100]); });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(se3), "Top inter-entités");
    downloadBlob(new Blob([XLSX.write(wb, { type: "array", bookType: "xlsx" })], { type: "application/octet-stream" }), "pomproc_analyse.xlsx");
  }

  return (
    <div><TopBar title="Analyse produits" onBack={back} right={<button className="pill" onClick={exportXlsx}>Exporter</button>} />
      <div className="scroll">
        <div className="chips">{[["all", "Tout"], ["90", "90 jours"], ["30", "30 jours"]].map(([k, l]) => <button key={k} className={"chip" + (range === k ? " chip-on" : "")} onClick={() => setRange(k)}>{l}</button>)}</div>
        <div className="kpi-row"><div className="kpi"><b>{A.count}</b><span>demandes</span></div><div className="kpi"><b>{A.top.length}</b><span>produits utilisés</span></div><div className="kpi"><b>{A.never.length}</b><span>jamais commandés</span></div></div>
        {A.count === 0 ? <div className="empty">Pas encore de données à analyser.</div> : (
          <div>
            <div className="chart-card">
              <div className="chart-h">Répartition de l'usage par catégorie</div>
              <div className="chart-sub">Part de chaque catégorie dans les articles commandés.</div>
              <div className="donut-wrap">
                <Donut segments={A.catShare.map((c) => ({ color: c.color, val: c.val }))} />
                <div className="donut-legend">{A.catShare.map((c) => <div key={c.catId} className="dl"><span className="an-dot" style={{ background: c.color }} /><span className="dl-name">{c.name}</span><span className="dl-val">{c.val}</span><span className="dl-pct">{Math.round(c.val / catTot * 100)}%</span></div>)}</div>
              </div>
            </div>

            <div className="chart-card">
              <div className="chart-h">Produits les plus utilisés</div>
              <div className="chips" style={{ margin: "6px 0 8px" }}>{[["freq", "Par demandes"], ["qty", "Par quantité"]].map(([k, l]) => <button key={k} className={"chip" + (sortBy === k ? " chip-on" : "")} onClick={() => setSortBy(k)}>{l}</button>)}</div>
              {rows.map((t, i) => <div key={t.pid} className="hb-row"><span className="hb-rank">{i + 1}</span><span className="hb-label" title={nm(t.pid)}>{nm(t.pid)}</span><div className="hb-track"><div className="hb-fill" style={{ width: Math.max(3, Math.round(t.val / maxVal * 100)) + "%", background: col(t.pid) }} /></div><span className="hb-num">{t.val}</span><span className="hb-cum">{t.cum}%</span></div>)}
              <div className="chart-sub" style={{ margin: "8px 0 0" }}>Colonne de droite = part cumulée (effet Pareto : repérez le petit nombre de références qui font l'essentiel).</div>
            </div>

            {A.heat.ids.length >= 2 && (
              <div className="chart-card">
                <div className="chart-h">Souvent commandés ensemble</div>
                <div className="chart-sub">Intensité de couleur = nombre de demandes où deux produits apparaissent ensemble (analyse de panier).</div>
                <div className="heat" style={{ gridTemplateColumns: "20px repeat(" + A.heat.ids.length + ", 1fr)" }}>
                  <div className="heat-corner" />
                  {A.heat.ids.map((_, j) => <div key={"c" + j} className="heat-hd">{j + 1}</div>)}
                  {A.heat.ids.map((pid, i) => (
                    <React.Fragment key={pid}>
                      <div className="heat-hd">{i + 1}</div>
                      {A.heat.ids.map((_, j) => { if (i === j) return <div key={j} className="heat-cell diag" />; const v = A.heat.matrix[i][j]; const a = v ? 0.15 + 0.85 * (v / A.heat.max) : 0; return <div key={j} className="heat-cell" style={{ background: v ? "rgba(18,138,82," + a + ")" : "#f3f7f4", color: a > 0.55 ? "#fff" : "#274b39" }}>{v || ""}</div>; })}
                    </React.Fragment>))}
                </div>
                <div className="heat-legend">{A.heat.ids.map((pid, i) => <div key={pid} className="hl"><span className="hl-n">{i + 1}</span><span className="an-dot" style={{ background: col(pid) }} /><span className="hl-name">{nm(pid)}</span></div>)}</div>
              </div>)}

            <div className="chart-card">
              <div className="chart-h">Le plus souvent remplacés / échangés</div>
              {A.replacedArr.length === 0 ? <div className="empty sm">Aucun remplacement enregistré.</div> : A.replacedArr.map((r) => { const mx = A.replacedArr[0].count || 1; return (<div key={r.pid}><div className="hb-row"><span className="hb-label" title={nm(r.pid)}>{nm(r.pid)}</span><div className="hb-track"><div className="hb-fill" style={{ width: Math.max(3, Math.round(r.count / mx * 100)) + "%", background: "#C9992B" }} /></div><span className="hb-num">{r.count}×</span></div>{r.top ? <div className="an-sub">→ souvent : {r.top}</div> : null}</div>); })}
            </div>

            <div className="chart-card">
              <div className="chart-h">Le plus souvent manquants</div>
              {A.missingArr.length === 0 ? <div className="empty sm">Aucune rupture enregistrée.</div> : A.missingArr.map((m) => { const mx = A.missingArr[0].c || 1; return (<div key={m.pid} className="hb-row"><span className="hb-label" title={nm(m.pid)}>{nm(m.pid)}</span><div className="hb-track"><div className="hb-fill" style={{ width: Math.max(3, Math.round(m.c / mx * 100)) + "%", background: "#D9453F" }} /></div><span className="hb-num">{m.c}×</span></div>); })}
            </div>

            <div className="chart-card">
              <div className="chart-h">Flux inter-entités (FFA · DC · IIS)</div>
              <div className="chart-sub">Origine / entité facturée des articles consommés sur les chantiers IIS. FFA et DC = achetés en externe.</div>
              {(() => { const maxc = Math.max(1, ...ENT_ORDER.map((e) => (A.entTot[e]?.cost || 0))); return ENT_ORDER.map((e) => { const v = A.entTot[e] || { qty: 0, cost: 0 }; return (
                <div key={e} className="hb-row"><span className="hb-label"><span className="ebadge" style={{ background: ENTITIES[e].color }}>{e}</span></span><div className="hb-track"><div className="hb-fill" style={{ width: Math.max(3, Math.round((v.cost || 0) / maxc * 100)) + "%", background: ENTITIES[e].color }} /></div><span className="hb-num">{money(v.cost)}</span><span className="hb-cum">{v.qty}u</span></div>); }); })()}
              <div className="flow-h">Qui prend chez qui — propriétaire → facturé à</div>
              <div className="flow-grid" style={{ gridTemplateColumns: "34px repeat(" + ENT_ORDER.length + ", 1fr)" }}>
                <div className="flow-corner" />
                {ENT_ORDER.map((c) => <div key={c} className="flow-hd"><span className="ebadge" style={{ background: ENTITIES[c].color }}>{c}</span></div>)}
                {ENT_ORDER.map((owner) => (<React.Fragment key={owner}><div className="flow-hd"><span className="ebadge" style={{ background: ENTITIES[owner].color }}>{owner}</span></div>{ENT_ORDER.map((billed) => { const f = A.flowArr.find((x) => x.owner === owner && x.billed === billed); const c = f?.cost || 0; return <div key={billed} className={"flow-cell" + (c ? "" : " zero") + (owner !== billed && c ? " xfer" : "")}>{c ? money(c) : "—"}</div>; })}</React.Fragment>))}
              </div>
              <div className="chart-sub" style={{ marginTop: 6 }}>Ligne = propriétaire du matériel · colonne = entité facturée. Les cases hors-diagonale = transferts entre entités.</div>
              <div className="flow-h">Top articles transférés (hors IIS)</div>
              {A.crossArr.length === 0 ? <div className="empty sm">Aucun transfert inter-entités enregistré.</div> : A.crossArr.map((c) => { const p = productById[c.pid] || extraById[c.pid]; const mx = A.crossArr[0].cost || 1; return (
                <div key={c.pid} className="hb-row"><span className="hb-label" title={p?.nom || c.pid}><span className="ebadge" style={{ background: ENTITIES[c.entity].color }}>{c.entity}</span> {p?.nom || c.pid}</span><div className="hb-track"><div className="hb-fill" style={{ width: Math.max(3, Math.round(c.cost / mx * 100)) + "%", background: ENTITIES[c.entity].color }} /></div><span className="hb-num">{money(c.cost)}</span><span className="hb-cum">{c.qty}u</span></div>); })}
            </div>

            <div className="chart-card">
              <div className="chart-h">Jamais commandés ({A.never.length})</div>
              <div className="chart-sub">Candidats à retirer du catalogue.</div>
              <div className="never-wrap">{A.never.slice(0, 30).map((p) => <span key={p.id} className="never-chip"><span className="an-dot" style={{ background: catById[p.catId]?.color || "#999" }} />{p.nom}</span>)}{A.never.length > 30 ? <span className="never-chip">+{A.never.length - 30}…</span> : null}</div>
            </div>
          </div>)}
        <p className="pane-note" style={{ marginTop: 6 }}>Les paires « souvent ensemble » sont d'excellentes bases de templates. L'export Excel reprend toutes ces analyses en détail (une feuille par thème).</p>
      </div>
    </div>
  );
}

/* ============================ Fiche / PDF ============================ */
function Sheet({ order, session, data, productById, catById, addItems, back }) {
  const pById = { ...productById }; (order.extras || []).forEach((e) => (pById[e.id] = e));
  const cats = [...data.productCats].sort((a, b) => a.order - b.order);
  const grouped = cats.map((c) => ({ cat: c, lines: order.lines.filter((l) => pById[l.productId]?.catId === c.id).sort((a, b) => byName(pById[a.productId] || { nom: "" }, pById[b.productId] || { nom: "" })) })).filter((g) => g.lines.length);
  const hasNum = grouped.some((g) => g.cat.num);
  const cols = hasNum ? 7 : 6;
  const done = order.status === "terminee";
  const lbk = {}; order.lines.forEach((l) => (lbk[lineKey(l)] = l));
  const issues = Object.entries(order.prep || {}).filter(([, v]) => v.state && v.state !== "ok");
  const canPrint = session?.role === "magasinier" || session?.role === "admin";
  const canAdd = session && (session.role === "magasinier" || session.role === "admin" || order.username === session.username);
  const addBatches = (() => { const m = {}; order.lines.forEach((l) => { if (l.addedAt) { const k = new Date(l.addedAt).toISOString().slice(0, 10); (m[k] = m[k] || { date: l.addedAt, by: l.addedBy, items: [] }).items.push(l); } }); return Object.values(m).sort((a, b) => a.date - b.date); })();
  const dOnly = (ms) => new Date(ms).toLocaleDateString("fr-BE");

  // Pagination A4
  const items = [];
  grouped.forEach((g) => { items.push({ type: "cat", cat: g.cat }); g.lines.forEach((l) => items.push({ type: "line", line: l, cat: g.cat })); });
  const CAP = 40, SIG = 8;
  const pages = []; let cur = [];
  items.forEach((it) => {
    if (cur.length >= CAP || (it.type === "cat" && cur.length >= CAP - 1)) { pages.push(cur); cur = []; }
    cur.push(it);
  });
  if (cur.length) pages.push(cur);
  if (!pages.length) pages.push([]);
  if (pages[pages.length - 1].length > CAP - SIG) pages.push([]);
  const nPages = pages.length;

  const copies = done
    ? [{ k: "admin", banner: "ADMINISTRATION", cls: "ban-admin" }, { k: "ret", banner: "DOCUMENT IMPORTANT À RAMENER", cls: "ban-ret" }]
    : [{ k: "work", banner: null, cls: "" }];

  const theadEl = (
    <thead>
      <tr className="hdr-brand"><td colSpan={cols}>
        <div className="hb-wrap">
          <span className="hb-left"><img src={LOGO} className="hb-logo" alt="XLG" /><b>{order.ref}</b></span>
          <span className="hb-info">
            <span className="hbi wrapok"><i>Chantier</i> {order.nom} · N° {order.numero}</span>
            <span className="hbi"><i>Demandeur</i> {order.username}</span>
            <span className="hbi"><i>Chef d'équipe</i> {order.responsable || "—"}</span>
            <span className="hbi"><i>Départ</i> {fmtD(order.departDate)}{order.departTime ? " " + order.departTime : ""}</span>
            <span className="hbi"><i>Demande du</i> {dOnly(order.createdAt)}</span>
          </span>
          <span className="hb-iis">IIS</span>
        </div>
      </td></tr>
      <tr className="hdr-cols"><th className="c-ref">Réf.</th><th>Désignation</th><th className="c-d">D</th><th className="c-lr">L</th><th className="c-lr">R</th>{hasNum && <th className="c-no">N°</th>}<th className="c-fac">Fact.</th></tr>
    </thead>
  );

  const renderRow = (l, key) => { const p = pById[l.productId]; const pr = order.prep?.[lineKey(l)]; const dLabel = l.mode === "paquet" ? l.qty + "×" + p.qteEmballage : l.qty + " " + p.unite.toLowerCase().slice(0, 3); const L = pr?.state === "missing" ? "0" : (pr?.qtyPrep != null ? pr.qtyPrep : (pr?.state ? l.qty : "")); const bill = ENTITIES[l.entity]?.bill;
    return (<tr key={key} className={pr && pr.state === "missing" ? "ln-missing" : ""}>
      <td className="c-ref">{p.ref}</td>
      <td className="c-name">{p.nom}{l.variant ? " — " + l.variant : ""}{p.custom ? " *" : ""}{l.addedAt ? <span className="add-tag">＋{dOnly(l.addedAt)}</span> : null}{pr?.state === "replaced" && pr.repl ? <em className="repl"> → {pr.repl}</em> : null}</td>
      <td className="c-d">{dLabel}</td>
      <td className="c-lr">{L !== "" ? L : <span className="wline" />}</td>
      <td className="c-lr"><span className="wline" /></td>
      {hasNum && <td className="c-no">{g_num(l)}</td>}
      <td className="c-fac">{bill ? <b style={{ color: ENTITIES[l.entity].color }}>{l.entity}</b> : <span className="wline" />}</td>
    </tr>); };
  function g_num(l) { const cat = catById[pById[l.productId]?.catId]; return cat?.num ? (l.assetNo || <span className="wline" />) : ""; }

  const afterEl = (
    <div className="doc-after">
      <div className="dlr-legend">D = Demandé · L = Livré · R = Retour</div>
      <div className="doc-bill">
        <div className="db-total">TOTAL <b>{money(order.total)}</b></div>
        {(order.billFFA > 0 || order.billDC > 0) && <div className="db-sub">à facturer — FFA : <b>{money(order.billFFA)}</b> · DC : <b>{money(order.billDC)}</b></div>}
        <div className="db-chk">Préparation terminée <span className="bigbox" /></div>
      </div>
      {(order.extras || []).length > 0 && <div className="db-sub" style={{ marginTop: 4 }}>* produit hors catalogue</div>}
      {addBatches.length > 0 && <div className="doc-adds"><b>Ajouts après création :</b>{addBatches.map((b, i) => <span key={i} className="add-batch"> · <u>{dOnly(b.date)}</u>{b.by ? " (" + b.by + ")" : ""} : {b.items.map((l) => { const p = pById[l.productId]; return (p?.nom || l.productId) + (l.variant ? " " + l.variant : "") + " ×" + l.qty; }).join(", ")}</span>)}</div>}
      {done && (issues.length > 0 || order.magasinierNote) && (
        <div className="doc-cr"><b>Compte-rendu ({order.completedBy}) :</b>{issues.map(([key, v]) => { const l = lbk[key]; const p = pById[l?.productId]; return <span key={key}> · {p?.nom || key}{l?.variant ? " " + l.variant : ""} : {v.state === "missing" ? "manquant" : "remplacé" + (v.repl ? " par " + v.repl : "")}</span>; })}{order.magasinierNote ? <div>Note : {order.magasinierNote}</div> : null}</div>)}
      <div className="doc-sign">
        <div className="ds"><span>Demandeur</span><div className="ds-name">{order.username}</div><div className="ds-line">Signature</div></div>
        <div className="ds"><span>Chef d'équipe</span><div className="ds-name">{order.responsable || ""}</div><div className="ds-line">Signature</div></div>
        <div className="ds"><span>Magasinier</span><div className="ds-name">{order.completedBy || ""}</div><div className="ds-line">Signature</div></div>
      </div>
    </div>
  );

  const labelEl = (
    <div className="doc-label"><div className="lab-inner">
      <div className="lab-top"><img src={LOGO} className="lab-logo" alt="XLG" /><span className="lab-iis">IIS</span><span className="lab-ref">{order.ref}</span></div>
      <div className="lab-resp-lbl">Responsable chantier</div>
      <div className="lab-resp">{order.responsable || "—"}</div>
      <div className="lab-sub">{order.nom} · N° {order.numero}</div>
      <div className="lab-depart">Départ {fmtD(order.departDate)}{order.departTime ? " · " + order.departTime : ""}</div>
      <div className="lab-foot"><span>Demandeur : {order.username}</span><span>Préparé par : {order.completedBy || "—"}</span></div>
    </div></div>
  );

  return (
    <div>
      <TopBar title={"Fiche " + order.ref} onBack={back} right={canPrint ? <button className="pill" onClick={() => window.print()}>Imprimer / PDF</button> : null} />
      <div className="scroll sheet-scroll">
        {canAdd && <button className="add-items-btn no-print" onClick={addItems}>＋ Ajouter des articles à cette demande</button>}
        <div id="print-area">
          {copies.map((cp) => (
            <div className="copy" key={cp.k}>
              {pages.map((pg, pi) => (
                <div className="a4page" key={pi}>
                  {pi === 0 && cp.banner && <div className={"copy-banner " + cp.cls}>{cp.banner}</div>}
                  <table className="doc-table">
                    {theadEl}
                    <tbody>
                      {pg.map((it, ii) => it.type === "cat"
                        ? <tr key={ii} className="cat-row"><td colSpan={cols} style={{ background: it.cat.color }}><span className="cr-in"><CatIcon icon={it.cat.icon} color="#fff" size={11} /> {it.cat.order}. {it.cat.name}</span></td></tr>
                        : renderRow(it.line, ii))}
                    </tbody>
                  </table>
                  {pi === nPages - 1 && afterEl}
                  <div className="page-foot">Page {pi + 1} / {nPages}{cp.banner ? " · " + cp.banner : ""}</div>
                </div>
              ))}
              {labelEl}
            </div>
          ))}
          <div className="a4page notes-page">
            <div className="notes-h"><img src={LOGO} className="hb-logo" alt="XLG" /> Notes / ajouts manuscrits — {order.ref}</div>
            <div className="notes-lines">{Array.from({ length: 24 }).map((_, i) => <div key={i} className="nl" />)}</div>
          </div>
        </div>
        {canPrint && <p className="print-hint no-print">Impression A4 : {done ? "2 copies (Administration + Document à ramener), " : ""}liste paginée avec numéros de page, puis l'étiquette responsable (tournez la feuille d'un quart de tour).</p>}
        {(order.prepPhotos || []).length > 0 && <div className="sheet-photos no-print"><div className="sec-h">Photos de la préparation</div><div className="photo-grid">{order.prepPhotos.map((src, i) => <div key={i} className="photo-thumb static"><img src={src} alt="" /></div>)}</div></div>}
      </div>
    </div>
  );
}/* ============================ Administration ============================ */
function Admin({ data, saveData, back }) {
  const [tab, setTab] = useState("users");
  function resetDemo() {
    if (!window.confirm("Recharger le catalogue de démonstration ?\n\nLes produits, catégories et templates actuels seront remplacés par les données d'exemple (avec les chaînages associés). Les comptes utilisateurs sont conservés.")) return;
    saveData({ ...data, productCats: SEED.productCategories, products: SEED.products, templates: SEED.templates, kits: SEED.kits || [] });
    window.alert("Catalogue de démonstration rechargé.");
  }
  return (
    <div><TopBar title="Administration" onBack={back} />
      <div className="admin-tabs no-print">{[["users", "Utilisateurs"], ["cats", "Catégories"], ["prod", "Produits"], ["kits", "Kits"]].map(([k, l]) => <button key={k} className={"atab" + (tab === k ? " atab-on" : "")} onClick={() => setTab(k)}>{l}</button>)}</div>
      <div className="scroll">
        {tab === "users" && <AdminUsers data={data} saveData={saveData} />}
        {tab === "cats" && <AdminCats data={data} saveData={saveData} />}
        {tab === "prod" && <AdminProd data={data} saveData={saveData} />}
        {tab === "kits" && <AdminKits data={data} saveData={saveData} />}
        <div className="admin-reset"><button className="reset-btn" onClick={resetDemo}>↺ Recharger le catalogue de démo</button><div className="reset-note">Remplace produits, catégories et templates par les données d'exemple (avec les associations Karcher → tuyau → raccord, générateur → essence). Les comptes sont conservés.</div></div>
      </div>
    </div>
  );
}
function AdminUsers({ data, saveData }) {
  const [f, setF] = useState({ username: "", password: "", workplace: "", role: "user", email: "", phone: "" });
  function add() { if (!f.username.trim() || !f.password || !f.workplace.trim()) return; if (data.users.some((u) => u.username.toLowerCase() === f.username.trim().toLowerCase())) return; saveData({ ...data, users: [...data.users, { ...f, username: f.username.trim(), workplace: f.workplace.trim(), email: f.email.trim(), phone: f.phone.trim() }] }); setF({ username: "", password: "", workplace: "", role: "user", email: "", phone: "" }); }
  const del = (u) => { if (u.username === "admin") return; saveData({ ...data, users: data.users.filter((x) => x !== u) }); };
  const upd = (u, patch) => saveData({ ...data, users: data.users.map((x) => (x === u ? { ...x, ...patch } : x)) });
  return (
    <div className="pane">
      <div className="card"><div className="card-h">Nouvel utilisateur</div>
        <input className="inp" placeholder="Identifiant" value={f.username} onChange={(e) => setF({ ...f, username: e.target.value })} autoCapitalize="none" />
        <input className="inp" placeholder="Mot de passe" value={f.password} onChange={(e) => setF({ ...f, password: e.target.value })} />
        <input className="inp" placeholder="Lieu de travail" value={f.workplace} onChange={(e) => setF({ ...f, workplace: e.target.value })} />
        <input className="inp" placeholder="E-mail (notifications)" value={f.email} onChange={(e) => setF({ ...f, email: e.target.value })} autoCapitalize="none" inputMode="email" />
        <input className="inp" placeholder="Téléphone WhatsApp (ex : 32470…)" value={f.phone} onChange={(e) => setF({ ...f, phone: e.target.value })} inputMode="tel" />
        <select className="inp" value={f.role} onChange={(e) => setF({ ...f, role: e.target.value })}>{ROLES.map((r) => <option key={r.id} value={r.id}>{r.label}</option>)}</select>
        <button className="btn btn-primary btn-block" onClick={add}>Ajouter</button></div>
      <div className="list">{data.users.map((u, i) => (
        <div key={i} className="row static u-row"><div className="row-main">
          <div className="row-title">{u.username} <span className="badge">{ROLES.find((r) => r.id === u.role)?.label}</span></div>
          <div className="row-sub">{u.workplace}</div>
          <div className="u-contact">
            <input className="inp mini" placeholder="E-mail" value={u.email || ""} onChange={(e) => upd(u, { email: e.target.value })} autoCapitalize="none" inputMode="email" />
            <input className="inp mini" placeholder="Tél. WhatsApp" value={u.phone || ""} onChange={(e) => upd(u, { phone: e.target.value })} inputMode="tel" />
          </div>
        </div>{u.username !== "admin" && <button className="del big" onClick={() => del(u)}>✕</button>}</div>))}</div>
    </div>
  );
}
function AdminCats({ data, saveData }) {
  const cats = [...data.productCats].sort((a, b) => a.order - b.order);
  const [f, setF] = useState({ name: "", order: "", num: false, color: "#128A52", icon: "mat", variants: "" });
  function add() { if (!f.name.trim()) return; saveData({ ...data, productCats: [...data.productCats, { id: uid("cat_"), name: f.name.trim(), order: parseInt(f.order, 10) || cats.length + 1, num: !!f.num, color: f.color, icon: f.icon, variants: parseGroups(f.variants) }] }); setF({ name: "", order: "", num: false, color: "#128A52", icon: "mat", variants: "" }); }
  const upd = (c, patch) => saveData({ ...data, productCats: data.productCats.map((x) => x.id === c.id ? { ...x, ...patch } : x) });
  function del(c) { if (data.products.some((p) => p.catId === c.id) && !window.confirm("Des produits seront supprimés. Continuer ?")) return; saveData({ ...data, productCats: data.productCats.filter((x) => x !== c), products: data.products.filter((p) => p.catId !== c.id) }); }
  return (
    <div className="pane">
      <div className="card"><div className="card-h">Nouvelle catégorie</div>
        <input className="inp" placeholder="Nom" value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} />
        <div className="row3b"><input className="inp" placeholder="Ordre" value={f.order} onChange={(e) => setF({ ...f, order: e.target.value })} inputMode="numeric" /><input className="inp cin" type="color" value={f.color} onChange={(e) => setF({ ...f, color: e.target.value })} /><select className="inp" value={f.icon} onChange={(e) => setF({ ...f, icon: e.target.value })}>{Object.keys(ICONS).map((k) => <option key={k} value={k}>{k}</option>)}</select></div>
        <textarea className="inp ta" placeholder="Variantes (une ligne par type) — ex :\nTaille: S, M, L, XL\nCouleur: Rouge, Bleu" value={f.variants} onChange={(e) => setF({ ...f, variants: e.target.value })} />
        <label className="chk"><input type="checkbox" checked={f.num} onChange={(e) => setF({ ...f, num: e.target.checked })} /> Demande un N° de matériel</label>
        <button className="btn btn-primary btn-block" onClick={add}>Ajouter</button></div>
      {cats.map((c) => (
        <div key={c.id} className="card cat-admin">
          <div className="cat-admin-top"><span className="cat-preview" style={{ background: c.color }}><CatIcon icon={c.icon} color="#fff" size={18} /></span><input className="inp cat-name-field" value={c.name} onChange={(e) => upd(c, { name: e.target.value })} placeholder="Nom de la catégorie" /><button className="del big" onClick={() => del(c)}>✕</button></div>
          <div className="cat-admin-row">
            <div className="caf"><span>Ordre</span><input className="inp" value={c.order} onChange={(e) => upd(c, { order: parseInt(e.target.value, 10) || c.order })} inputMode="numeric" /></div>
            <div className="caf"><span>Couleur</span><input className="inp cin" type="color" value={c.color} onChange={(e) => upd(c, { color: e.target.value })} /></div>
            <div className="caf"><span>Icône</span><select className="inp" value={c.icon} onChange={(e) => upd(c, { icon: e.target.value })}>{Object.keys(ICONS).map((k) => <option key={k} value={k}>{k}</option>)}</select></div>
            <label className="chk sm caf-num"><input type="checkbox" checked={!!c.num} onChange={() => upd(c, { num: !c.num })} /> N°</label>
          </div>
          <textarea className="inp catvar ta" defaultValue={groupsToText(c.variants)} onBlur={(e) => upd(c, { variants: parseGroups(e.target.value) })} placeholder="Variantes — ex : Taille: S, M, L (nouvelle ligne) Couleur: Rouge, Bleu" />
        </div>))}
    </div>
  );
}function AdminProd({ data, saveData }) {
  const cats = [...data.productCats].sort((a, b) => a.order - b.order);
  const fileRef = useRef();
  const [cat, setCat] = useState(cats[0]?.id || "");
  const [edit, setEdit] = useState(null);
  const [q, setQ] = useState("");
  const blank = { ref: "", nom: "", unite: "Pièce", qteEmballage: "1", prix: "", entity: "IIS", image: "", variants: "", assoc: [] };
  const [f, setF] = useState(blank);
  const [aq, setAq] = useState("");
  const [newAssoc, setNewAssoc] = useState(null);
  const [uploading, setUploading] = useState(false);
  const reset = () => { setEdit(null); setF(blank); setAq(""); };
  async function onPickImage(e) { const file = e.target.files?.[0]; if (!file) return; try { setUploading(true); const url = await uploadImage(file); setF((v) => ({ ...v, image: url })); } catch (err) { window.alert("Image impossible : " + (err?.message || err)); } finally { setUploading(false); e.target.value = ""; } }
  function save() {
    if (!f.nom.trim() || !cat) return;
    const prod = { catId: cat, ref: f.ref.trim(), nom: f.nom.trim(), unite: f.unite.trim() || "Pièce", qteEmballage: parseInt(f.qteEmballage, 10) || 1, prix: parseFloat(String(f.prix).replace(",", ".")) || 0, entity: f.entity, image: f.image.trim(), variants: parseGroups(f.variants), assoc: f.assoc };
    if (edit) saveData({ ...data, products: data.products.map((p) => p.id === edit ? { ...p, ...prod } : p) }); else saveData({ ...data, products: [...data.products, { id: uid("p"), ...prod }] });
    reset();
  }
  function startEdit(p) { setEdit(p.id); setCat(p.catId); setF({ ref: p.ref, nom: p.nom, unite: p.unite, qteEmballage: String(p.qteEmballage), prix: String(p.prix).replace(".", ","), entity: p.entity || "IIS", image: p.image || "", variants: groupsToText(p.variants), assoc: p.assoc || [] }); }
  const del = (p) => saveData({ ...data, products: data.products.filter((x) => x !== p) });
  const list = (q.trim() ? data.products.filter((p) => p.nom.toLowerCase().includes(q.trim().toLowerCase()) || String(p.ref).toLowerCase().includes(q.trim().toLowerCase())) : data.products.filter((p) => p.catId === cat)).sort(byName);
  const assocResults = aq.trim() ? data.products.filter((p) => (p.nom.toLowerCase().includes(aq.trim().toLowerCase()) || String(p.ref).toLowerCase().includes(aq.trim().toLowerCase())) && p.id !== edit && !f.assoc.includes(p.id)).slice(0, 6) : [];
  function createAssocProduct() { if (!newAssoc.nom.trim()) return; const prod = { id: uid("p"), catId: newAssoc.catId, ref: "", nom: newAssoc.nom.trim(), unite: newAssoc.unite.trim() || "Pièce", qteEmballage: 1, prix: parseFloat(String(newAssoc.prix).replace(",", ".")) || 0, entity: newAssoc.entity, image: "", variants: [], assoc: [] }; saveData({ ...data, products: [...data.products, prod] }); setF((ff) => ({ ...ff, assoc: [...ff.assoc, prod.id] })); setNewAssoc(null); setAq(""); }
  function exportTemplate() {
    const rows = [["Catégorie", "Réf", "Description", "Unité", "Qté emballage", "Prix", "Entité", "Image (URL)"]];
    [...data.products].sort((a, b) => { const ca = cats.find((c) => c.id === a.catId), cb = cats.find((c) => c.id === b.catId); return (ca?.order || 0) - (cb?.order || 0) || byName(a, b); }).forEach((p) => { const c = cats.find((x) => x.id === p.catId); rows.push([c?.name || "", p.ref, p.nom, p.unite, p.qteEmballage, p.prix, p.entity || "IIS", p.image || ""]); });
    const ws = XLSX.utils.aoa_to_sheet(rows); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Produits");
    downloadBlob(new Blob([XLSX.write(wb, { type: "array", bookType: "xlsx" })], { type: "application/octet-stream" }), "pomproc_produits.xlsx");
  }
  async function onImport(e) {
    const file = e.target.files?.[0]; if (!file) return;
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
      const catByName = {}; data.productCats.forEach((c) => (catByName[c.name.trim().toLowerCase()] = c.id));
      const byRef = {}; data.products.forEach((p) => (byRef[String(p.ref).trim()] = p));
      const out = []; let skipped = 0;
      rows.slice(1).forEach((r) => { if (!r || (!r[1] && !r[2])) return; const catId = catByName[String(r[0] || "").trim().toLowerCase()]; if (!catId) { skipped++; return; } const ref = String(r[1] ?? "").trim(); const ex = byRef[ref]; out.push({ id: ex?.id || uid("p"), catId, ref, nom: String(r[2] ?? "").trim(), unite: String(r[3] ?? "Pièce").trim() || "Pièce", qteEmballage: parseInt(r[4], 10) || 1, prix: parseFloat(String(r[5] ?? "0").replace(",", ".")) || 0, entity: ENTITIES[String(r[6] || "IIS").trim().toUpperCase()] ? String(r[6]).trim().toUpperCase() : "IIS", image: String(r[7] ?? "").trim(), variants: ex?.variants || [], assoc: ex?.assoc || [] }); });
      if (!out.length) return window.alert("Aucun produit valide (vérifiez la colonne Catégorie).");
      if (!window.confirm("Remplacer le catalogue par " + out.length + " produit(s) ?" + (skipped ? " (" + skipped + " ligne(s) ignorée(s))" : ""))) return;
      await saveData({ ...data, products: out }); window.alert("Catalogue mis à jour : " + out.length + " produit(s).");
    } catch { window.alert("Fichier illisible."); }
    e.target.value = "";
  }
  return (
    <div className="pane">
      <div className="card"><div className="card-h">Import / export Excel</div>
        <button className="btn btn-ghost btn-block" onClick={exportTemplate}>⭳ Télécharger le modèle (tous les produits)</button>
        <button className="btn btn-primary btn-block" onClick={() => fileRef.current?.click()}>⭱ Importer un fichier Excel</button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={onImport} />
        <p className="pane-note" style={{ margin: "10px 0 0" }}>Modifiez le fichier (catégorie par son nom) puis réimportez : le catalogue est remplacé (variantes et associations conservées par référence).</p>
      </div>
      <div className="card"><div className="card-h">{edit ? "Modifier le produit" : "Ajouter un produit"}</div>
        <select className="inp" value={cat} onChange={(e) => setCat(e.target.value)}>{cats.map((c) => <option key={c.id} value={c.id}>{c.order}. {c.name}</option>)}</select>
        <input className="inp" placeholder="Description" value={f.nom} onChange={(e) => setF({ ...f, nom: e.target.value })} />
        <input className="inp" placeholder="Référence" value={f.ref} onChange={(e) => setF({ ...f, ref: e.target.value })} />
        <div className="row3"><input className="inp" placeholder="Unité" value={f.unite} onChange={(e) => setF({ ...f, unite: e.target.value })} /><input className="inp" placeholder="Qté/emb." value={f.qteEmballage} onChange={(e) => setF({ ...f, qteEmballage: e.target.value })} inputMode="numeric" /><input className="inp" placeholder="Prix" value={f.prix} onChange={(e) => setF({ ...f, prix: e.target.value })} inputMode="decimal" /></div>
        <div className="img-row"><input className="inp" placeholder="URL de la photo (ou upload →)" value={f.image} onChange={(e) => setF({ ...f, image: e.target.value })} /><label className="up-btn">{uploading ? "…" : "Photo"}<input type="file" accept="image/*" style={{ display: "none" }} onChange={onPickImage} /></label></div>
        {f.image ? <img className="img-prev" src={f.image} alt="" /> : null}
        <label className="lbl">Variantes (facultatif — vide = celles de la catégorie)</label><textarea className="inp ta" placeholder="Une ligne par type — ex :\nTaille: S, M, L\nCouleur: Rouge, Bleu" value={f.variants} onChange={(e) => setF({ ...f, variants: e.target.value })} />
        <label className="lbl">Produits associés (à prévoir avec)</label>
        {f.assoc.length > 0 && <div className="assoc-edit">{f.assoc.map((aid) => { const ap = data.products.find((x) => x.id === aid); return <span key={aid} className="assoc-chip has">{ap?.nom || aid}<button onClick={() => setF({ ...f, assoc: f.assoc.filter((x) => x !== aid) })}>✕</button></span>; })}</div>}
        <input className="inp" placeholder="Rechercher un produit à associer…" value={aq} onChange={(e) => setAq(e.target.value)} />
        {aq.trim() && (
          <div className="assoc-res">
            {assocResults.map((p) => <button key={p.id} className="assoc-opt" onClick={() => { setF((ff) => ({ ...ff, assoc: [...ff.assoc, p.id] })); setAq(""); }}>{p.nom} <small>{p.ref}</small></button>)}
            {assocResults.length === 0 && <div className="assoc-none">Aucun produit existant ne correspond.</div>}
            <button className="assoc-opt assoc-new" onClick={() => setNewAssoc({ nom: aq.trim(), catId: cat, unite: "Pièce", prix: "", entity: "IIS" })}>+ Créer un nouveau produit « {aq.trim()} »</button>
          </div>)}
        {newAssoc && (
          <div className="new-assoc">
            <div className="na-h">Nouveau produit associé</div>
            <input className="inp" placeholder="Description" value={newAssoc.nom} onChange={(e) => setNewAssoc({ ...newAssoc, nom: e.target.value })} />
            <select className="inp" value={newAssoc.catId} onChange={(e) => setNewAssoc({ ...newAssoc, catId: e.target.value })}>{cats.map((c) => <option key={c.id} value={c.id}>{c.order}. {c.name}</option>)}</select>
            <div className="row3"><input className="inp" placeholder="Unité" value={newAssoc.unite} onChange={(e) => setNewAssoc({ ...newAssoc, unite: e.target.value })} /><input className="inp" placeholder="Prix" value={newAssoc.prix} onChange={(e) => setNewAssoc({ ...newAssoc, prix: e.target.value })} inputMode="decimal" /><select className="inp" value={newAssoc.entity} onChange={(e) => setNewAssoc({ ...newAssoc, entity: e.target.value })}>{ENT_ORDER.map((en) => <option key={en} value={en}>{en}</option>)}</select></div>
            <button className="btn btn-primary btn-block" onClick={createAssocProduct}>Créer et associer</button>
            <button className="btn btn-ghost btn-block" onClick={() => setNewAssoc(null)}>Annuler</button>
          </div>)}
        <label className="lbl">Entité propriétaire</label>
        <div className="ent-seg full">{ENT_ORDER.map((e) => <button key={e} className={"ent-b" + (f.entity === e ? " ent-on" : "")} style={f.entity === e ? { background: ENTITIES[e].color, borderColor: ENTITIES[e].color } : {}} onClick={() => setF({ ...f, entity: e })}>{e}{!ENTITIES[e].bill ? " (interne)" : ""}</button>)}</div>
        <button className="btn btn-primary btn-block" onClick={save}>{edit ? "Enregistrer" : "Ajouter"}</button>
        {edit && <button className="btn btn-ghost btn-block" onClick={reset}>Annuler</button>}
      </div>
      <div className="search"><span>⌕</span><input placeholder="Rechercher un produit à modifier…" value={q} onChange={(e) => setQ(e.target.value)} />{q && <button onClick={() => setQ("")}>✕</button>}</div>
      <div className="catwrap">{cats.map((c) => <button key={c.id} className="ccat" style={cat === c.id ? { background: c.color, borderColor: c.color, color: "#fff" } : { color: c.color, borderColor: c.color }} onClick={() => setCat(c.id)}><CatIcon icon={c.icon} color={cat === c.id ? "#fff" : c.color} size={14} /><span>{c.order}. {c.name}</span></button>)}</div>
      <div className="list">{list.map((p) => (
        <div key={p.id} className="row static">
          <div className="prod-thumb sm">{p.image ? <img src={p.image} alt="" /> : <CatIcon icon={cats.find((c) => c.id === p.catId)?.icon} color={cats.find((c) => c.id === p.catId)?.color} size={18} />}</div>
          <div className="row-main"><div className="row-title">{p.nom} <EntBadge e={p.entity} /></div><div className="row-sub">{q.trim() ? (cats.find((c) => c.id === p.catId)?.name || "?") + " · " : ""}Réf {p.ref} · {p.unite} · emb. {p.qteEmballage}{p.prix ? " · " + money(p.prix) : ""}{p.variants && p.variants.length ? " · variantes" : ""}{p.assoc && p.assoc.length ? " · " + p.assoc.length + " assoc." : ""}</div></div>
          <button className="mini-btn" onClick={() => startEdit(p)}>✎</button><button className="del big" onClick={() => del(p)}>✕</button>
        </div>))}</div>
    </div>
  );
}

/* ============================ Admin : Kits ============================ */
function AdminKits({ data, saveData }) {
  const products = data.products || [];
  const pById = useMemo(() => { const m = {}; products.forEach((p) => (m[p.id] = p)); return m; }, [products]);
  const catById = useMemo(() => { const m = {}; (data.productCats || []).forEach((c) => (m[c.id] = c)); return m; }, [data.productCats]);
  const kits = data.kits || [];
  const [edit, setEdit] = useState(null);
  const [q, setQ] = useState("");
  function newKit() { setEdit({ id: uid("kit_"), name: "", items: [] }); setQ(""); }
  function startEdit(k) { setEdit(JSON.parse(JSON.stringify(k))); setQ(""); }
  function save() { if (!edit.name.trim() || !edit.items.length) { window.alert("Nom et au moins un article requis."); return; } const exists = kits.some((k) => k.id === edit.id); const next = exists ? kits.map((k) => (k.id === edit.id ? edit : k)) : [...kits, edit]; saveData({ ...data, kits: next }); setEdit(null); }
  function del(k) { if (!window.confirm("Supprimer le kit « " + k.name + " » ?")) return; saveData({ ...data, kits: kits.filter((x) => x.id !== k.id) }); }
  function addItem(p) { setEdit((e) => ({ ...e, items: [...e.items, { productId: p.id, variant: "", qty: 1, mode: "unite", entity: p.entity || "IIS" }] })); setQ(""); }
  const updItem = (i, patch) => setEdit((e) => ({ ...e, items: e.items.map((it, idx) => (idx === i ? { ...it, ...patch } : it)) }));
  const delItem = (i) => setEdit((e) => ({ ...e, items: e.items.filter((_, idx) => idx !== i) }));
  const results = q.trim() ? products.filter((p) => p.nom.toLowerCase().includes(q.trim().toLowerCase()) || String(p.ref).toLowerCase().includes(q.trim().toLowerCase())).slice(0, 8) : [];
  if (edit) {
    return (
      <div className="pane">
        <div className="prod-head2"><h2 className="pane-h">{kits.some((k) => k.id === edit.id) ? "Modifier le kit" : "Nouveau kit"}</h2></div>
        <label className="lbl">Nom du kit</label><input className="inp" value={edit.name} onChange={(e) => setEdit({ ...edit, name: e.target.value })} placeholder="Ex : Kit démarrage chantier" />
        <label className="lbl">Articles du kit ({edit.items.length})</label>
        <div className="list">{edit.items.length === 0 && <div className="empty sm">Aucun article. Ajoutez-en ci-dessous.</div>}
          {edit.items.map((it, i) => { const p = pById[it.productId]; const cat = catById[p?.catId]; const groups = variantGroupsOf(p, catById); return (
            <div key={i} className="kit-item">
              <div className="kit-item-top">{cat ? <CatTag cat={cat} /> : null}<span className="kit-item-name">{p?.nom || it.productId}</span><button className="del" onClick={() => delItem(i)}>✕</button></div>
              <div className="kit-item-ctrl">
                <Stepper value={it.qty} onChange={(v) => updItem(i, { qty: Math.max(1, v) })} />
                {p?.qteEmballage > 1 && <div className="seg mini"><button className={"seg-b" + (it.mode !== "paquet" ? " seg-on" : "")} onClick={() => updItem(i, { mode: "unite" })}>Unité</button><button className={"seg-b" + (it.mode === "paquet" ? " seg-on" : "")} onClick={() => updItem(i, { mode: "paquet" })}>×{p.qteEmballage}</button></div>}
                {groups.length === 1 && <select className="inp mini" value={it.variant} onChange={(e) => updItem(i, { variant: e.target.value })}><option value="">Variante…</option>{groups[0].options.map((o) => <option key={o} value={o}>{o}</option>)}</select>}
                <div className="ent-seg mini">{ENT_ORDER.map((en) => <button key={en} className={"ent-b" + (it.entity === en ? " ent-on" : "")} style={it.entity === en ? { background: ENTITIES[en].color, borderColor: ENTITIES[en].color } : {}} onClick={() => updItem(i, { entity: en })}>{en}</button>)}</div>
              </div>
            </div>); })}
        </div>
        <label className="lbl">Ajouter un article</label>
        <div className="search"><span>⌕</span><input placeholder="Rechercher un produit ou une réf…" value={q} onChange={(e) => setQ(e.target.value)} />{q && <button onClick={() => setQ("")}>✕</button>}</div>
        <div className="assoc-res">{results.map((p) => <button key={p.id} className="assoc-opt" onClick={() => addItem(p)}>{p.nom} <small>{p.ref}</small></button>)}</div>
        <button className="btn btn-primary btn-block" onClick={save}>Enregistrer le kit</button>
        <button className="btn btn-ghost btn-block" onClick={() => setEdit(null)}>Annuler</button>
      </div>);
  }
  return (
    <div className="pane">
      <button className="btn btn-primary btn-block" onClick={newKit}>+ Nouveau kit</button>
      <p className="pane-note">Les kits sont des lots prédéfinis d'articles que le demandeur peut ajouter d'un coup dans une demande.</p>
      <div className="list">{kits.length === 0 ? <div className="empty">Aucun kit défini.</div> : kits.map((k) => (
        <div key={k.id} className="row static kit-row">
          <button className="kit-open" onClick={() => startEdit(k)}>
            <div className="row-title">🎁 {k.name}</div>
            <div className="row-sub">{k.items.length} article(s)</div>
          </button>
          <button className="del big" onClick={() => del(k)}>✕</button>
        </div>))}</div>
    </div>);
}

/* ================================ CSS ================================ */
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap');
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
:root{--g900:#06301f;--g700:#0b5e3a;--g:#128A52;--g500:#1aa863;--g300:#7dd6a6;--bg:#eef8f1;--line:#cfe6d8;--ink:#0c1a12;--muted:#5f7a6b;--dc:#D9453F;--dc:#D9453F}
.app{font-family:"Montserrat","Helvetica Neue",Arial,sans-serif;color:var(--ink);background:var(--bg);max-width:520px;margin:0 auto;min-height:100vh;display:flex;flex-direction:column}
.center{min-height:100vh;display:flex;align-items:center;justify-content:center}
.spinner{width:34px;height:34px;border:3px solid var(--line);border-top-color:var(--g);border-radius:50%;animation:spin .8s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.scroll{flex:1;overflow-y:auto;padding:16px;padding-bottom:28px;scrollbar-gutter:stable}
.pb-bar{padding-bottom:100px}
.topbar{display:flex;align-items:center;gap:10px;background:var(--g900);color:#fff;padding:12px 14px;position:sticky;top:0;z-index:5;min-height:56px}
.topbar-logo{height:22px}.icon-btn{background:rgba(255,255,255,.16);border:0;color:#fff;width:32px;height:32px;border-radius:9px;font-size:22px;cursor:pointer}
.topbar-title{font-weight:700;font-size:16px;flex:1}.topbar-right{margin-left:auto}
.pill{background:#fff;color:var(--g900);border:0;border-radius:999px;padding:7px 13px;font-weight:700;font-size:13px;cursor:pointer;font-family:inherit}
.pill-out{background:transparent;color:#fff;border:1px solid rgba(255,255,255,.45)}
.stepdots{display:flex;gap:6px}.dot{width:8px;height:8px;border-radius:50%;background:rgba(255,255,255,.3)}.dot.on{background:var(--g300)}.dot.done{background:var(--g500)}

.auth{padding:0 18px 24px;flex:1;display:flex;flex-direction:column}
.auth-hero{text-align:center;padding:40px 0 18px}.auth-logo{height:42px;margin-bottom:16px}
.auth-app{font-weight:800;font-size:34px;color:var(--g900)}.auth-app span{color:var(--g500)}
.auth-iis{color:var(--muted);font-weight:500;font-size:14px}.auth-iis b{color:var(--g)}
.auth-card{margin-top:6px}.auth-note{text-align:center;color:var(--muted);font-size:12px;margin-top:14px}
.card{background:#fff;border:1px solid var(--line);border-radius:16px;padding:16px;box-shadow:0 1px 2px rgba(6,48,31,.05);margin-bottom:12px}
.card-h{font-weight:700;margin-bottom:10px;color:var(--g900)}
.lbl{display:block;font-size:12.5px;font-weight:600;color:var(--muted);margin:12px 0 5px}
.inp{width:100%;border:1px solid var(--line);border-radius:11px;padding:12px 13px;font-size:15px;font-family:inherit;background:#fff;color:var(--ink)}
.inp:focus{outline:none;border-color:var(--g);box-shadow:0 0 0 3px rgba(18,138,82,.14)}
.ta{min-height:70px;resize:vertical}
.cin{padding:4px;height:44px}
.row2b{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.row3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-top:8px}
.row3b{display:grid;grid-template-columns:1fr 64px 1fr;gap:8px;margin-top:8px;align-items:center}
.chk{display:flex;align-items:center;gap:6px;font-size:13px;font-weight:600;color:var(--g900);margin-top:10px}
.chk.sm{font-size:12px;margin:0}.chk input{width:17px;height:17px;accent-color:var(--g)}
.err{color:var(--dc);font-size:13px;margin-top:10px;font-weight:600}
.btn{border:0;border-radius:12px;padding:13px 16px;font-size:15px;font-weight:700;font-family:inherit;cursor:pointer}
.btn-primary{background:var(--g);color:#fff}.btn-primary:disabled{background:#a9cdba}
.btn-ghost{background:transparent;color:var(--g);border:1px solid var(--line)}
.btn-block{width:100%;margin-top:12px}

.hello{margin:2px 0 12px}.hello-name{font-weight:800;font-size:20px;color:var(--g900)}.hello-sub{color:var(--muted);font-weight:500;font-size:13px}
.notif{background:#e6f7ee;border:1px solid var(--g300);color:var(--g700);border-radius:12px;padding:11px 13px;font-size:13px;font-weight:600;margin-bottom:12px;cursor:pointer}
.cta{width:100%;display:flex;align-items:center;gap:14px;background:var(--g900);color:#fff;border:0;border-radius:16px;padding:16px;cursor:pointer;text-align:left;font-family:inherit}
.cta-plus{width:42px;height:42px;flex:none;border-radius:12px;background:var(--g500);display:flex;align-items:center;justify-content:center;font-size:26px}
.cta b{display:block;font-size:16px}.cta small{display:block;color:rgba(255,255,255,.75);font-weight:500}
.quick{display:grid;grid-template-columns:1fr 1fr;gap:9px;margin-top:12px}
.qbtn{position:relative;display:flex;align-items:center;gap:9px;background:#fff;border:1px solid var(--line);border-radius:13px;padding:14px;font-weight:700;color:var(--g900);font-family:inherit;font-size:14px;cursor:pointer;text-align:left}
.qi{width:28px;height:28px;flex:none;border-radius:8px;background:var(--bg);color:var(--g);display:flex;align-items:center;justify-content:center;font-size:15px}
.qi-badge{position:absolute;top:8px;right:8px;background:var(--dc);color:#fff;border-radius:999px;font-size:11px;font-style:normal;font-weight:800;padding:1px 7px}
.sec-h{font-weight:800;color:var(--g900);margin:20px 0 4px;font-size:15px}
.empty{text-align:center;color:var(--muted);padding:30px 20px;line-height:1.6}.empty.sm{padding:16px}
.list{display:flex;flex-direction:column;gap:9px;margin-top:12px}
.row{display:flex;align-items:center;gap:10px;background:#fff;border:1px solid var(--line);border-radius:13px;padding:12px;cursor:pointer;text-align:left;font-family:inherit;width:100%}
.row.static{cursor:default}.row-main{flex:1;min-width:0}
.row-title{font-weight:700;color:var(--g900);font-size:15px}.row-sub{color:var(--muted);font-size:12.5px;margin-top:2px}
.row-right{display:flex;align-items:center;gap:6px}.chev{color:var(--g500);font-size:20px}
.ref-tag{font-weight:700;color:var(--muted);font-size:12px}
.badge{background:var(--g);color:#fff;font-size:10.5px;padding:2px 7px;border-radius:999px}
.ebadge{color:#fff;font-size:10px;font-weight:800;padding:1px 6px;border-radius:5px;letter-spacing:.3px}
.stbadge{color:#fff;font-size:10.5px;font-weight:800;padding:2px 9px;border-radius:999px}
.cattag{display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:8px;flex:none}

.pane{padding-bottom:8px}.pane-h{font-size:19px;font-weight:800;color:var(--g900);margin:2px 0 8px}
.pane-note{color:var(--muted);font-size:12.5px;line-height:1.5}
.demandeur-note{margin-top:16px;color:var(--muted);font-size:12.5px}.demandeur-note b{color:var(--g900)}
.search{display:flex;align-items:center;gap:8px;background:#fff;border:1px solid var(--line);border-radius:12px;padding:0 12px;margin-bottom:12px}
.search span{color:var(--muted);font-size:17px}.search input{flex:1;border:0;padding:12px 0;font-size:15px;font-family:inherit;background:transparent}
.search input:focus{outline:none}.search button{border:0;background:none;color:var(--muted);font-size:15px;cursor:pointer}
.chips{display:flex;gap:8px;overflow-x:auto;padding-bottom:4px;margin-bottom:12px}
.chip{white-space:nowrap;border:1px solid var(--line);background:#fff;color:var(--g900);border-radius:999px;padding:8px 14px;font-weight:600;font-size:13px;cursor:pointer;font-family:inherit}
.chip-on{background:var(--g900);color:#fff;border-color:var(--g900)}
.catwrap{display:flex;flex-wrap:wrap;gap:7px;margin-bottom:12px}
.ccat{display:flex;align-items:center;gap:6px;white-space:nowrap;border:1.5px solid;background:#fff;border-radius:10px;padding:7px 11px;font-weight:700;font-size:12.5px;cursor:pointer;font-family:inherit}
.prod-head2{display:flex;align-items:center;gap:10px;min-height:30px;margin-bottom:8px}
.backcat{background:var(--bg);border:1px solid var(--line);color:var(--g900);border-radius:9px;padding:7px 12px;font-weight:700;font-size:13px;cursor:pointer;font-family:inherit;flex:none}
.cur-cat{display:inline-flex;align-items:center;gap:7px;font-weight:800;color:var(--g900);font-size:16px}
.catgrid{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.catbig{position:relative;display:flex;flex-direction:column;align-items:flex-start;gap:8px;background:#fff;border:1.5px solid;border-radius:14px;padding:14px;cursor:pointer;font-family:inherit;text-align:left;min-height:98px}
.catbig-ic{width:44px;height:44px;border-radius:12px;display:flex;align-items:center;justify-content:center}
.catbig-name{font-weight:800;color:var(--g900);font-size:14px;line-height:1.2;padding-right:22px}
.catbig-badge{position:absolute;top:12px;right:12px;color:#fff;font-weight:800;font-size:12px;min-width:22px;height:22px;border-radius:999px;display:flex;align-items:center;justify-content:center;padding:0 6px}
.catbig-go{position:absolute;bottom:10px;right:12px;color:var(--muted);font-size:20px;font-weight:700}
.prodlist{display:flex;flex-direction:column;gap:9px}
.prod{background:#fff;border:1px solid var(--line);border-radius:13px;padding:11px}
.prod-top{display:flex;align-items:center;gap:11px}
.prod-thumb{width:40px;height:40px;flex:none;border:1.5px solid var(--line);border-radius:10px;display:flex;align-items:center;justify-content:center;overflow:hidden;background:#fff}
.prod-thumb img{width:100%;height:100%;object-fit:cover}.prod-thumb.sm{width:34px;height:34px;border-radius:9px}
.prod-info{flex:1;min-width:0}.prod-name{font-weight:600;font-size:14px;line-height:1.3}.prod-meta{color:var(--muted);font-size:11.5px;margin-top:3px}
.prod-opts{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-top:11px;padding-top:11px;border-top:1px dashed var(--line);flex-wrap:wrap}
.seg{display:flex;border:1px solid var(--line);border-radius:9px;overflow:hidden}
.seg-b{border:0;background:#fff;padding:7px 9px;font-size:12px;font-weight:600;color:var(--muted);cursor:pointer;font-family:inherit}
.seg-b.seg-on{background:var(--g);color:#fff}.seg-b:disabled{opacity:.4}
.ent-seg{display:flex;gap:5px}.ent-seg.full{display:grid;grid-template-columns:1fr 1fr 1fr;margin-top:4px}
.ent-b{border:1px solid var(--line);background:#fff;color:var(--muted);border-radius:8px;padding:7px 10px;font-size:12px;font-weight:800;cursor:pointer;font-family:inherit}
.ent-b.ent-on{color:#fff}
.asset-inp{width:100%;margin-top:9px;border:1px dashed var(--g300);border-radius:9px;padding:9px 11px;font-size:13px;font-family:inherit;background:var(--bg)}
.asset-inp:focus{outline:none;border-color:var(--g)}
.stepper{display:flex;align-items:center;border:1px solid var(--line);border-radius:10px;overflow:hidden;flex:none}
.stepper button{width:34px;height:34px;border:0;background:var(--bg);color:var(--g900);font-size:19px;font-weight:700;cursor:pointer}
.stepper input{width:36px;height:34px;border:0;text-align:center;font-size:15px;font-weight:700;font-family:inherit;color:var(--g900);border-left:1px solid var(--line);border-right:1px solid var(--line)}
.stepper input:focus{outline:none}
.bottombar{position:fixed;bottom:0;left:0;right:0;max-width:520px;margin:0 auto;background:#fff;border-top:1px solid var(--line);padding:12px 16px;display:flex;align-items:center;gap:12px;z-index:6}
.bb-total{flex:1;display:flex;flex-direction:column;line-height:1.1}.bb-total small{color:var(--muted);font-size:11px;font-weight:600}.bb-total b{font-size:18px;color:var(--g900)}
.bottombar .btn{min-width:150px}
.recap-box{background:#fff;border:1px solid var(--line);border-radius:13px;padding:6px 14px;margin-bottom:14px}
.recap-line{display:flex;justify-content:space-between;padding:9px 0;border-bottom:1px solid var(--line);font-size:14px}.recap-line:last-child{border-bottom:0}
.recap-line span{color:var(--muted)}.recap-line b{color:var(--g900);text-align:right}
.recap-items{display:flex;flex-direction:column;gap:8px}
.ritem{display:flex;align-items:center;gap:10px;background:#fff;border:1px solid var(--line);border-radius:12px;padding:9px 11px}
.ritem-main{flex:1;min-width:0}.ritem-name{font-weight:600;font-size:13.5px}.ritem-sub{color:var(--muted);font-size:12px;margin-top:2px}
.asset-tag{background:var(--bg);border:1px solid var(--g300);color:var(--g700);font-size:10.5px;padding:1px 6px;border-radius:5px}
.del{background:none;border:0;color:var(--dc);font-size:15px;cursor:pointer;font-weight:700}.del.big{font-size:18px;padding:6px}
.recap-total{display:flex;justify-content:space-between;align-items:center;background:var(--g900);color:#fff;border-radius:13px;padding:14px;margin-top:14px}.recap-total b{font-size:20px}
.mini-btn{background:var(--bg);border:1px solid var(--line);color:var(--g900);border-radius:9px;padding:7px 10px;font-weight:700;font-size:13px;cursor:pointer;font-family:inherit}

/* magasin board */
.board{display:flex;flex-direction:column;gap:10px}
.bcard{background:#fff;border:1px solid var(--line);border-left:5px solid;border-radius:13px;padding:13px;width:100%;text-align:left;font-family:inherit;cursor:pointer}
.bcard-top{display:flex;justify-content:space-between;align-items:center;margin-bottom:4px}.bref{font-weight:800;color:var(--g900);font-size:15px}
.bcard-name{font-weight:700;color:var(--ink);margin-bottom:8px}
.bcard-grid{display:grid;grid-template-columns:1fr 1fr;gap:7px}
.bcard-grid>div span{display:block;font-size:10px;color:var(--muted);text-transform:uppercase;font-weight:700;letter-spacing:.3px}
.bcard-grid>div b{font-size:13px;color:var(--g900)}
.bcard-depart b{color:var(--dc)}

/* prepare */
.prep-head{background:#fff;border:1px solid var(--line);border-radius:13px;padding:13px;margin-bottom:12px}
.prep-h-row{font-size:16px;color:var(--g900)}.prep-h-sub{color:var(--muted);font-size:12.5px;margin:3px 0 8px}
.prep-cat{margin:14px 0 4px}
.prep-cat-h{display:flex;align-items:center;gap:7px;color:#fff;font-weight:800;font-size:13px;padding:7px 11px;border-radius:10px 10px 0 0}
.prep-line{background:#fff;border:1px solid var(--line);border-top:0;padding:10px 11px}
.prep-line:last-child{border-radius:0 0 10px 10px}
.prep-name{font-weight:600;font-size:13.5px}.prep-sub{color:var(--muted);font-size:11.5px;margin-top:2px}
.prep-states{display:flex;gap:6px;margin-top:9px}
.pstate{flex:1;border:1px solid var(--line);background:#fff;border-radius:8px;padding:7px 4px;font-size:12px;font-weight:700;cursor:pointer;font-family:inherit}
.prep-edit{display:flex;flex-wrap:wrap;gap:12px;margin-top:9px;padding-top:9px;border-top:1px dashed var(--line);align-items:flex-end}
.pe-field{display:flex;flex-direction:column;gap:4px}
.pe-field>span{font-size:10px;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:.3px}
.asset-inp.mini{margin-top:0;width:150px}
.pstate:disabled,.ent-b:disabled{opacity:.5;cursor:not-allowed}
.inp:disabled,.ta:disabled,.asset-inp:disabled{background:#f3f7f4;color:#8b9a90}
.lock-banner{background:#fff7ec;border:1px solid #f0d9b0;color:#7a5a1e;border-radius:11px;padding:10px 12px;font-size:12.5px;font-weight:600;margin-bottom:12px}
.lock-banner.editable{background:#e6f7ee;border-color:var(--g300);color:var(--g700)}
.lock-note{flex:1;color:var(--muted);font-weight:700;font-size:13px;text-align:center}
.lbl-opt{font-weight:500;color:var(--muted);opacity:.8}
.u-row .u-contact{display:flex;gap:6px;margin-top:6px}
.inp.mini{padding:7px 9px;font-size:12.5px;margin:0}
.notify-card{background:#fff;border:1px solid var(--line);border-radius:14px;padding:12px 14px;margin:12px 16px 0}
.notify-h{font-size:13px;font-weight:800;color:var(--g900);margin-bottom:9px}
.notify-row{display:flex;flex-wrap:wrap;gap:8px}
.ntf-btn{display:inline-flex;align-items:center;gap:6px;padding:9px 14px;border-radius:999px;font-weight:700;font-size:13px;text-decoration:none;color:#fff}
.ntf-btn.wa{background:#25D366}.ntf-btn.mail{background:#2D6CDF}
.notify-empty{font-size:12.5px;color:var(--muted);line-height:1.4}
.filterbar{position:sticky;top:0;z-index:5;background:var(--bg);padding:6px 0 8px;margin:0 0 4px}
.chips{display:flex;gap:7px;overflow-x:auto;padding:8px 2px 2px;-webkit-overflow-scrolling:touch;scrollbar-width:none}
.chips::-webkit-scrollbar{display:none}
.chip{flex:none;white-space:nowrap;padding:8px 13px;border-radius:999px;border:1.5px solid var(--line);background:#fff;font-size:13px;font-weight:700;color:var(--g900);cursor:pointer}
.chip.on{background:var(--g900);border-color:var(--g900);color:#fff}
.qty-badge{position:absolute;top:-6px;right:-6px;min-width:20px;height:20px;padding:0 5px;border-radius:999px;color:#fff;font-size:11px;font-weight:800;display:flex;align-items:center;justify-content:center;box-shadow:0 0 0 2px #fff}
.prod-thumb{position:relative}
.row-opts{display:flex;flex-direction:column;gap:8px;margin-top:9px}
.details-btn{align-self:flex-start;background:var(--bg);border:1px solid var(--line);border-radius:9px;padding:7px 11px;font-size:12.5px;font-weight:700;color:var(--g900);cursor:pointer}
.assoc.inline{margin-top:0}
.basket-btn{position:relative;background:var(--bg);border:1.5px solid var(--line);border-radius:12px;padding:10px 14px 10px 12px;font-size:14px;font-weight:800;color:var(--g900);cursor:pointer;display:inline-flex;align-items:center;gap:7px}
.bb-count{background:var(--g);color:#fff;border-radius:999px;min-width:20px;height:20px;padding:0 5px;font-size:11.5px;font-weight:800;display:inline-flex;align-items:center;justify-content:center}
.basket-list{max-height:52vh;overflow-y:auto;margin:4px 0 8px}
.basket-row{display:flex;align-items:center;gap:10px;padding:9px 2px;border-bottom:1px solid var(--line)}
.basket-main{flex:1;min-width:0}
.basket-name{font-size:13.5px;font-weight:700;color:var(--g900);line-height:1.3}
.basket-sub{font-size:11.5px;color:var(--muted);margin-top:1px}
.admin-reset{margin:22px 16px 8px;padding-top:16px;border-top:1px dashed var(--line)}
.reset-btn{width:100%;background:#fff;border:1.5px solid var(--dc);color:var(--dc);border-radius:12px;padding:11px;font-weight:800;font-size:14px;cursor:pointer;font-family:inherit}
.reset-note{font-size:11.5px;color:var(--muted);line-height:1.4;margin-top:8px}

.bill-card{background:#fff;border:1px solid var(--line);border-radius:14px;padding:14px;margin-bottom:10px}
.bill-month{font-weight:800;color:var(--g900);text-transform:capitalize;display:flex;justify-content:space-between;align-items:baseline}.bill-month small{color:var(--muted);font-weight:500;font-size:12px}
.bill-rows{margin-top:10px;display:flex;flex-direction:column;gap:7px}.bill-r{display:flex;align-items:center;justify-content:space-between;font-size:14px}.bill-r b{color:var(--g900);font-weight:800}
.bill-r.total{border-top:1px solid var(--line);padding-top:8px}.bill-r.total span{color:var(--muted)}
.bill-r.delta{border-top:1px dashed var(--line);padding-top:6px;font-size:12px;align-items:flex-start}
.bill-r.delta span{color:var(--muted)}.bill-r.delta b{color:var(--dc);font-weight:700;text-align:right}
.bill-tot{color:var(--g900);font-weight:800}
.kpi-row{display:flex;gap:9px;margin:2px 0 8px}
.kpi{flex:1;background:#fff;border:1px solid var(--line);border-radius:12px;padding:11px 6px;text-align:center}
.kpi b{display:block;font-size:20px;color:var(--g900)}.kpi span{font-size:10.5px;color:var(--muted);font-weight:600}
.an-sec{margin-top:16px}
.an-h{font-weight:800;color:var(--g900);font-size:15px;margin-bottom:6px}
.an-note{color:var(--muted);font-size:11.5px;margin-bottom:8px;line-height:1.4}
.an-item{padding:7px 0;border-bottom:1px solid var(--line)}
.an-line{display:flex;align-items:center;gap:8px}
.an-rank{width:18px;text-align:right;font-weight:800;color:var(--muted);font-size:12px;flex:none}
.an-dot{width:10px;height:10px;border-radius:3px;flex:none}
.an-name{flex:1;min-width:0;font-weight:600;font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.an-val{font-weight:800;color:var(--g900);font-size:13px;flex:none;min-width:40px;text-align:right}
.an-bar-wrap{height:6px;margin-top:5px;margin-left:26px;background:var(--bg);border-radius:5px;overflow:hidden}
.an-bar{height:100%;border-radius:5px}
.an-pair{display:flex;align-items:center;gap:6px;padding:8px 0;border-bottom:1px solid var(--line);font-size:13px;flex-wrap:wrap}
.an-pair b{color:var(--g900)}.an-plus{color:var(--muted);font-weight:800}
.an-pcount{margin-left:auto;font-weight:800;color:var(--g900)}
.an-sub{color:var(--muted);font-size:11.5px;margin-top:2px}
.never-wrap{display:flex;flex-wrap:wrap;gap:6px}
.never-chip{display:inline-flex;align-items:center;gap:5px;background:#fff;border:1px solid var(--line);border-radius:999px;padding:5px 10px;font-size:12px;color:var(--g900)}
.chart-card{background:#fff;border:1px solid var(--line);border-radius:14px;padding:14px;margin-bottom:12px}
.chart-h{font-weight:800;color:var(--g900);font-size:14px}
.chart-sub{color:var(--muted);font-size:11.5px;margin-top:2px;line-height:1.4}
.donut-wrap{display:flex;align-items:center;gap:14px;margin-top:10px}
.donut-legend{flex:1;display:flex;flex-direction:column;gap:5px;min-width:0}
.dl{display:flex;align-items:center;gap:7px;font-size:12px}
.dl-name{flex:1;min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:var(--ink)}
.dl-val{font-weight:800;color:var(--g900)}.dl-pct{color:var(--muted);font-size:11px;width:36px;text-align:right}
.hb-row{display:flex;align-items:center;gap:8px;margin:7px 0}
.hb-rank{width:16px;text-align:right;font-weight:800;color:var(--muted);font-size:11px;flex:none}
.hb-label{width:33%;flex:none;font-size:12px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.hb-track{flex:1;height:16px;background:var(--bg);border-radius:6px;overflow:hidden}
.hb-fill{height:100%;border-radius:6px;min-width:3px}
.hb-num{flex:none;font-weight:800;color:var(--g900);font-size:12px;min-width:26px;text-align:right}
.hb-cum{flex:none;color:var(--muted);font-size:10px;width:30px;text-align:right}
.heat{display:grid;gap:3px;margin-top:10px}
.heat-hd{display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:800;color:var(--muted);min-height:16px}
.heat-cell{aspect-ratio:1;min-height:24px;display:flex;align-items:center;justify-content:center;border-radius:4px;font-size:11px;font-weight:800}
.heat-cell.diag{background:repeating-linear-gradient(45deg,#eef2f0,#eef2f0 3px,#f7faf8 3px,#f7faf8 6px)}
.heat-legend{margin-top:10px;display:flex;flex-direction:column;gap:4px}
.hl{display:flex;align-items:center;gap:6px;font-size:11.5px}
.hl-n{width:16px;height:16px;border-radius:4px;background:var(--bg);color:var(--g900);font-weight:800;font-size:10px;display:flex;align-items:center;justify-content:center;flex:none}
.hl-name{flex:1;min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.hc-btn{margin-left:auto}
.var-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-top:11px;padding-top:11px;border-top:1px dashed var(--line)}
.var-cell{display:flex;flex-direction:column;align-items:center;gap:5px;background:var(--bg);border:1px solid var(--line);border-radius:10px;padding:7px 4px}
.var-cell.on{border-color:var(--g);background:#eaf7ee}
.var-lbl{font-weight:800;color:var(--g900);font-size:12px}
.var-multi{margin-top:11px;padding-top:11px;border-top:1px dashed var(--line)}
.vg{margin-bottom:8px}
.vg-name{display:block;font-size:10px;color:var(--muted);text-transform:uppercase;font-weight:700;margin-bottom:4px}
.vg-opts{display:flex;flex-wrap:wrap;gap:6px}
.vg-opt{border:1px solid var(--line);background:#fff;color:var(--g900);border-radius:999px;padding:6px 11px;font-size:12px;font-weight:700;cursor:pointer;font-family:inherit}
.vg-opt.on{background:var(--g);border-color:var(--g);color:#fff}
.vg-add{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-top:8px;background:var(--bg);border-radius:10px;padding:8px 10px}
.vg-combo{font-weight:800;color:var(--g900);font-size:13px}
.vg-list{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px}
.vg-chip{display:inline-flex;align-items:center;gap:6px;background:#eaf7ee;border:1px solid var(--g);color:var(--g700);border-radius:999px;padding:4px 10px;font-size:12px;font-weight:700}
.vg-chip button{border:0;background:none;color:var(--dc);cursor:pointer;font-weight:800;padding:0}
.catvar.ta{min-height:52px}
.assoc{margin-top:10px;padding-top:10px;border-top:1px dashed var(--line);font-size:11.5px;color:var(--muted);display:flex;flex-wrap:wrap;gap:6px;align-items:center}
.assoc-chip{border:1px solid var(--g300);background:#fff;color:var(--g700);border-radius:999px;padding:5px 10px;font-size:12px;font-weight:700;cursor:pointer;font-family:inherit}
.assoc-chip.has{background:#eaf7ee;border-color:var(--g);display:inline-flex;align-items:center;gap:5px;cursor:default}
.assoc-chip.has button{border:0;background:none;color:var(--dc);cursor:pointer;font-weight:800;padding:0}
.extras-box{background:#fff;border:1px dashed var(--g300);border-radius:12px;padding:10px;margin-bottom:12px}
.extras-h{font-weight:800;color:var(--g900);font-size:12.5px;margin-bottom:4px}
.extra-row{display:flex;align-items:center;gap:8px;padding:6px 0;border-top:1px solid var(--line)}
.extra-row:first-of-type{border-top:0}
.extra-name{font-weight:600;font-size:13px;flex:1;min-width:0}
.extra-meta{color:var(--muted);font-size:11px;white-space:nowrap}
.assoc-edit{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:6px}
.assoc-res{display:flex;flex-direction:column;gap:4px;margin-top:6px}
.assoc-opt{text-align:left;background:var(--bg);border:1px solid var(--line);border-radius:9px;padding:8px 11px;font-family:inherit;font-size:13px;color:var(--g900);cursor:pointer}
.assoc-opt small{color:var(--muted)}
.assoc-new{color:var(--g);font-weight:800;border-style:dashed;border-color:var(--g300)}
.assoc-none{color:var(--muted);font-size:12px;padding:6px 2px}
.new-assoc{background:var(--bg);border:1px solid var(--line);border-radius:12px;padding:12px;margin-top:8px}
.na-h{font-weight:800;color:var(--g900);font-size:13px;margin-bottom:8px}
.catvar{margin-top:6px;font-size:13px}
.cat-admin{margin-bottom:9px}
.modal-bg{position:fixed;inset:0;background:rgba(6,48,31,.45);display:flex;align-items:flex-end;justify-content:center;z-index:20}
.modal{background:#fff;width:100%;max-width:520px;border-radius:18px 18px 0 0;padding:18px;max-height:86vh;overflow-y:auto}
.modal-h{font-weight:800;color:var(--g900);font-size:17px;margin-bottom:8px}
.admin-tabs{display:flex;gap:4px;background:var(--g900);padding:0 8px 8px;overflow-x:auto}
.atab{white-space:nowrap;border:0;background:rgba(255,255,255,.14);color:#fff;border-radius:9px;padding:8px 12px;font-weight:600;font-size:12.5px;cursor:pointer;font-family:inherit}
.atab-on{background:#fff;color:var(--g900)}
.ord-inp{width:34px;text-align:center;border:1px solid var(--line);border-radius:8px;padding:7px 0;font-weight:800;color:var(--g900);font-family:inherit}
.cat-name-inp{width:100%;border:0;border-bottom:1px dashed var(--line);font-weight:700;color:var(--g900);font-size:14px;font-family:inherit;padding:2px 0;background:transparent}
.cat-name-inp:focus{outline:none;border-color:var(--g)}

/* Document */
.sheet-scroll{background:#dfeee6}
.doc{background:#fff;border-radius:8px;padding:7mm;box-shadow:0 2px 10px rgba(6,48,31,.12);color:#111;font-size:11px}
.doc-head{display:flex;align-items:center;gap:10px;border-bottom:2px solid var(--g900);padding-bottom:5px;margin-bottom:6px}
.doc-logo{height:20px}.doc-ref{font-weight:800;color:var(--g900);font-size:16px;flex:1}.doc-iis{font-weight:800;color:var(--g);font-size:15px;letter-spacing:1px}
.doc-band{display:flex;flex-wrap:wrap;gap:3px 14px;border:1px solid var(--line);border-radius:6px;padding:5px 8px;margin-bottom:7px;font-size:10.5px}
.doc-band span{color:var(--g900);font-weight:700}.doc-band i{color:var(--muted);font-weight:600;font-style:normal;margin-right:4px;text-transform:uppercase;font-size:8px;letter-spacing:.3px}
.doc-band .dep{color:var(--dc)}
.doc-band .chk{display:flex;align-items:center;gap:4px}.bigbox{display:inline-block;width:12px;height:12px;border:1.4px solid #333;border-radius:2px}
.doc-table{width:100%;border-collapse:collapse}
.doc-table th{background:var(--bg);color:var(--g900);text-align:left;padding:3px 6px;border-bottom:1.5px solid var(--g900);font-size:8.5px;text-transform:uppercase}
.doc-table td{padding:2px 6px;border-bottom:1px solid #e7f0ea;font-size:9.8px;line-height:1.1}
.cat-row td{color:#fff;font-weight:800;padding:2.5px 6px}.cr-in{display:inline-flex;align-items:center;gap:5px;text-transform:uppercase;font-size:9px;letter-spacing:.4px}
.c-ok{width:20px;text-align:center;font-weight:800}.box{display:inline-block;width:11px;height:11px;border:1.3px solid #333;border-radius:2px}
.c-ref{width:50px;color:var(--muted);font-size:9.5px}.c-name{font-weight:500}.repl{color:var(--dc);font-style:italic}
.c-qty{width:52px;white-space:nowrap;font-weight:700;color:var(--g900)}.c-no{width:50px}.c-fac{width:34px;text-align:center;font-weight:800;color:var(--dc)}
.wline{display:inline-block;width:100%;min-width:26px;border-bottom:1px solid #bbb;height:11px}
.ln-missing .c-name{text-decoration:line-through;color:#999}
.doc-bill{display:flex;align-items:baseline;justify-content:space-between;margin-top:7px;padding-top:5px;border-top:2px solid var(--g900)}
.db-total{font-weight:800;color:var(--g900);font-size:13px}.db-total b{font-size:15px;margin-left:6px}.db-sub{font-size:10px;color:var(--muted)}.db-sub b{color:var(--g900)}
.doc-cr{margin-top:7px;padding:6px 8px;background:#fff7ec;border:1px solid #f0d9b0;border-radius:6px;font-size:9.5px;color:#7a5a1e}
.doc-sign{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-top:10px}
.ds{border:1px solid var(--line);border-radius:6px;padding:5px 7px;min-height:58px;display:flex;flex-direction:column}
.ds span{font-size:7.5px;color:var(--muted);text-transform:uppercase;font-weight:700}.ds-name{font-weight:700;color:var(--g900);margin-top:2px;font-size:11px;line-height:1.15;word-break:break-word}
.ds-line{margin-top:auto;border-top:1px solid var(--line);padding-top:2px;font-size:7.5px;color:var(--muted)}
.doc-foot{display:flex;justify-content:space-between;margin-top:7px;padding-top:4px;border-top:1px solid var(--line);font-size:8px;color:var(--muted)}
.print-hint{color:var(--muted);font-size:12px;text-align:center;margin:12px 4px 0}
.photo-grid{display:flex;flex-wrap:wrap;gap:8px;margin-top:6px}
.photo-thumb{position:relative;width:84px;height:84px;border-radius:10px;overflow:hidden;border:1px solid var(--line)}
.photo-thumb img{width:100%;height:100%;object-fit:cover}
.photo-thumb button{position:absolute;top:3px;right:3px;background:rgba(0,0,0,.55);color:#fff;border:0;border-radius:6px;width:20px;height:20px;font-size:12px;cursor:pointer;font-weight:800;line-height:1}
.photo-thumb.static{width:104px;height:104px}
.photo-add{display:flex;align-items:center;justify-content:center;width:84px;height:84px;border:1.5px dashed var(--g300);border-radius:10px;color:var(--g);font-weight:700;font-size:13px;cursor:pointer}
.sheet-photos{margin-top:14px}

/* magasin compact */
.mag-list{display:flex;flex-direction:column;gap:6px}
.mag-row{display:flex;align-items:center;gap:10px;background:#fff;border:1px solid var(--line);border-left:4px solid;border-radius:10px;padding:9px 11px;width:100%;text-align:left;font-family:inherit;cursor:pointer}
.mag-ref{font-weight:800;color:var(--g900);font-size:13px;width:62px;flex:none}
.mag-mid{flex:1;min-width:0}
.mag-name{font-weight:700;color:var(--ink);font-size:13.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.mag-sub{color:var(--muted);font-size:11px;margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.mag-dep{text-align:right;flex:none;line-height:1.05}
.mag-dep span{display:block;font-size:8px;color:var(--muted);text-transform:uppercase;font-weight:700}
.mag-dep b{font-size:12px;color:var(--dc)}.mag-dep .hh{display:block;font-size:11px;color:var(--g900)}
/* upload image */
.img-row{display:flex;gap:8px;align-items:stretch}
.up-btn{display:flex;align-items:center;white-space:nowrap;background:var(--g);color:#fff;border-radius:11px;padding:0 14px;font-weight:700;font-size:13px;cursor:pointer}
.img-prev{margin-top:8px;width:70px;height:70px;object-fit:cover;border-radius:10px;border:1px solid var(--line)}
/* document header répété + étiquette */
.hdr-brand td{background:var(--g900);padding:5px 8px}
.hb-wrap{display:flex;align-items:center;gap:10px}
.hb-left{display:inline-flex;align-items:center;gap:8px;flex:none}.hb-logo{height:15px}.hb-left b{color:#fff;font-size:14px;font-weight:800}
.hb-info{flex:1;display:flex;flex-wrap:wrap;gap:1px 12px;color:#eaf7ef;font-size:9.5px;line-height:1.35}
.hb-info .hbi{white-space:nowrap}.hb-info .hbi.wrapok{white-space:normal}
.hb-info i{color:#9fe3bf;font-style:normal;font-weight:700;text-transform:uppercase;margin-right:3px}
.hb-iis{flex:none;color:var(--g300);font-weight:800;font-size:12px;letter-spacing:1px}
.hdr-cols th{background:var(--bg);color:var(--g900);text-align:left;padding:2.5px 6px;border-bottom:1.5px solid var(--g900);font-size:8px;text-transform:uppercase}
.doc-after{margin-top:6px}
.db-chk{font-size:10px;color:var(--g900);font-weight:700;display:flex;align-items:center;gap:5px}
.doc-label{width:100%;max-width:520px;margin:0 auto 14px;background:#fff;box-shadow:0 2px 12px rgba(6,48,31,.14);border-radius:3px;aspect-ratio:297/210;display:flex;align-items:center;justify-content:center;padding:16px;overflow:hidden}
.lab-inner{background:#fff;border:8px solid var(--g);border-radius:14px;padding:20px 26px;text-align:center;box-shadow:none;width:100%}
.lab-top{display:flex;align-items:center;justify-content:center;gap:12px;margin-bottom:12px}
.lab-logo{height:22px}.lab-iis{color:var(--g);font-weight:800;font-size:22px;letter-spacing:3px}.lab-ref{color:var(--muted);font-weight:700}
.lab-resp-lbl{font-size:13px;color:var(--muted);text-transform:uppercase;letter-spacing:1px;font-weight:700;margin-top:4px}
.lab-resp{font-size:54px;font-weight:800;color:var(--g900);line-height:1.03;margin:4px 0 16px;word-break:break-word}
.lab-sub{font-size:26px;font-weight:800;color:var(--g);line-height:1.15}
.lab-depart{font-size:18px;color:var(--muted);font-weight:600;margin-top:10px}


.version{text-align:center;color:var(--muted);font-size:11px;margin:22px 0 4px;opacity:.85}
.prep-line.treated{border-color:var(--g300)}
.prep-qty{display:flex;align-items:center;gap:10px;margin-top:9px;flex-wrap:wrap}
.prep-qty>span:first-child{font-size:11px;color:var(--muted);text-transform:uppercase;font-weight:700}
.qty-of{color:var(--muted);font-size:11.5px}
.repl-row{display:flex;gap:8px;margin-top:9px}
.price-inp{max-width:140px}
.prep-total{display:flex;justify-content:space-between;align-items:center;background:var(--bg);border-radius:10px;padding:10px 12px;margin:14px 0 4px;font-weight:700;color:var(--g900)}
.prep-total b{font-size:16px}
.c-d{width:44px;text-align:center;white-space:nowrap;font-weight:700;color:var(--g900);font-size:9.5px}
.c-lr{width:26px;text-align:center;font-weight:700;color:var(--g900)}
.dlr-legend{font-size:9px;color:var(--muted);margin-top:4px}
.lab-foot{display:flex;justify-content:space-between;gap:16px;margin-top:18px;font-size:12px;color:var(--muted);width:100%}
.cat-admin{padding:12px}
.cat-admin-top{display:flex;align-items:center;gap:8px;margin-bottom:8px}
.cat-preview{width:34px;height:34px;border-radius:9px;display:flex;align-items:center;justify-content:center;flex:none}
.cat-name-field{font-weight:700}
.cat-admin-row{display:flex;gap:8px;align-items:flex-end;flex-wrap:wrap}
.caf{display:flex;flex-direction:column;gap:3px}
.caf>span{font-size:10px;color:var(--muted);font-weight:700;text-transform:uppercase}
.caf .inp{padding:8px}
.caf .cin{height:38px;padding:3px}
.caf-num{align-self:center;margin-left:auto}
.cat-admin .catvar{margin-top:8px}


.sheet-scroll{padding-top:14px}
.a4page{background:#fff;width:100%;max-width:520px;aspect-ratio:210/297;margin:0 auto 14px;padding:16px;box-shadow:0 2px 12px rgba(6,48,31,.14);color:#111;font-size:10px;position:relative;overflow:visible;border-radius:3px}
.copy-banner{font-weight:800;text-align:center;letter-spacing:1px;padding:6px 8px;border-radius:6px;margin-bottom:8px;color:#fff;font-size:13px;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.ban-admin{background:var(--g900)}
.ban-ret{background:var(--dc)}
.page-foot{margin-top:8px;text-align:right;font-size:9px;color:var(--muted)}
.a4page .doc-after{margin-top:8px}


.vg-chip2{display:flex;align-items:center;gap:8px;background:#eaf7ee;border:1px solid var(--g);border-radius:10px;padding:5px 8px;font-size:12px;color:var(--g900);flex-wrap:wrap;margin-bottom:6px}
.vg-chip2 b{color:var(--g900)}
.vg-ent{display:flex;gap:4px;margin-left:auto}
.vg-eb{border:1px solid var(--line);background:#fff;border-radius:7px;padding:3px 7px;font-size:10.5px;font-weight:800;cursor:pointer;font-family:inherit}
.vg-x{border:0;background:none;color:var(--dc);font-weight:800;cursor:pointer;padding:0 2px}
.repl-box{margin-top:9px;display:flex;flex-direction:column;gap:6px}
.repl-search{display:flex;align-items:center;gap:6px;background:var(--bg);border:1px dashed var(--g300);border-radius:9px;padding:0 10px}
.repl-search span{color:var(--muted)}
.repl-search input{flex:1;border:0;background:transparent;padding:9px 0;font-family:inherit;font-size:13px;color:var(--ink)}
.repl-search input:focus{outline:none}
.repl-res{display:flex;flex-direction:column;gap:4px}
.repl-res button{text-align:left;background:#fff;border:1px solid var(--line);border-radius:8px;padding:7px 10px;font-family:inherit;font-size:12.5px;color:var(--g900);cursor:pointer}
.repl-res small{color:var(--muted)}

@media print{
  @page { size: A4 portrait; margin: 7mm; }
  body{background:#fff}.app{max-width:none}.no-print{display:none !important}
  .scroll{overflow:visible;padding:0}.sheet-scroll{background:#fff}
  .doc{box-shadow:none;border-radius:0;padding:0;font-size:9.3px}
  thead{display:table-header-group}
  tr{break-inside:avoid;page-break-inside:avoid}
  .doc-table td{padding:1.3px 5px;font-size:8.4px}
  .cat-row td{padding:1.8px 5px}
  .hdr-brand td,.hdr-cols th,.cat-row td,.ebadge,.hb-iis,.doc-cr,.stbadge,.doc-label{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  .doc-after{break-inside:avoid}
  .a4page{max-width:none;aspect-ratio:auto;width:100%;padding:0;margin:0;box-shadow:none;border-radius:0;font-size:8.6px;break-after:page}
  .copy + .copy{break-before:page}
  .copy-banner,.c-fac b{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  .doc-label{position:relative;height:283mm;margin:0;max-width:none;aspect-ratio:auto;box-shadow:none;padding:0;display:block}
  .lab-inner{position:absolute;top:50%;left:50%;width:283mm;height:194mm;margin:0;padding:16mm;border-width:10px;border-radius:0;box-shadow:none;display:flex;flex-direction:column;justify-content:center;transform:translate(-50%,-50%) rotate(-90deg);transform-origin:center}
  .lab-resp{font-size:104px;margin:10px 0 26px}.lab-sub{font-size:48px}.lab-depart{font-size:28px}.lab-resp-lbl{font-size:20px}.lab-logo{height:36px}.lab-iis{font-size:36px}.lab-foot{font-size:20px;margin-top:26px}
}
/* ---- Facturation : total & paiements ---- */
.grand-card{background:linear-gradient(135deg,var(--g900),var(--g700));color:#fff;border-radius:16px;padding:16px;margin-bottom:14px}
.grand-h{font-size:12.5px;font-weight:700;opacity:.85}
.grand-total{font-size:30px;font-weight:800;margin:2px 0 6px;letter-spacing:-.5px}
.grand-split{font-size:13px;opacity:.95;display:flex;align-items:center;gap:2px;flex-wrap:wrap}
.grand-split .ebadge{margin-right:3px}
.grand-pay{display:flex;gap:10px;margin-top:12px}
.gp{flex:1;background:rgba(255,255,255,.14);border-radius:11px;padding:9px 11px;display:flex;flex-direction:column}
.gp span{font-size:11px;opacity:.85;font-weight:600}
.gp b{font-size:16px;font-weight:800}
.gp.unpaid b{color:#ffd7d4}
.bill-r.paid-r b{color:var(--g500)}
.bill-r.unpaid-r b{color:var(--dc)}
.pay-row{display:flex;align-items:center;gap:8px;padding:9px 2px;border-bottom:1px solid var(--line)}
.pay-main{flex:1;min-width:0;text-align:left;background:none;border:0;font-family:inherit;cursor:pointer;padding:0}
.pay-toggle{flex:none;border:1.5px solid var(--dc);color:var(--dc);background:#fff;border-radius:999px;padding:7px 13px;font-size:12.5px;font-weight:800;cursor:pointer;font-family:inherit}
.pay-toggle.on{border-color:var(--g);color:#fff;background:var(--g)}
.pay-badge{flex:none;border-radius:999px;padding:6px 12px;font-size:12px;font-weight:800;background:#fdecea;color:var(--dc)}
.pay-badge.on{background:#eaf7ee;color:var(--g700)}
/* ---- Analyse : flux inter-entités ---- */
.flow-h{font-size:13px;font-weight:800;color:var(--g900);margin:16px 0 8px}
.flow-grid{display:grid;gap:4px;align-items:stretch}
.flow-corner{}
.flow-hd{display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700}
.flow-cell{display:flex;align-items:center;justify-content:center;background:#f3f7f4;border-radius:7px;padding:8px 4px;font-size:11px;font-weight:700;color:var(--g900);min-height:34px;text-align:center}
.flow-cell.zero{color:var(--muted);opacity:.5;font-weight:500}
.flow-cell.xfer{background:#fff4e6;color:#8a5a00;box-shadow:inset 0 0 0 1.5px #f0c47a}
/* ---- Kits ---- */
.kit-card{background:#fff;border:1.5px solid var(--line);border-radius:14px;padding:12px 14px;margin-bottom:10px}
.kit-card-h{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:8px}
.kit-card-name{font-size:15px;font-weight:800;color:var(--g900)}
.kit-add{border:0;background:var(--g);color:#fff;border-radius:999px;padding:8px 15px;font-size:13px;font-weight:800;cursor:pointer;font-family:inherit;flex:none}
.kit-card-items{display:flex;flex-wrap:wrap;gap:6px}
.kit-chip{background:var(--bg);border:1px solid var(--line);border-radius:8px;padding:4px 9px;font-size:12px;color:var(--g900);font-weight:600}
.kit-row .kit-open{flex:1;text-align:left;background:none;border:0;font-family:inherit;cursor:pointer;padding:0}
.kit-item{border:1px solid var(--line);border-radius:12px;padding:10px 12px;margin-bottom:8px;background:#fff}
.kit-item-top{display:flex;align-items:center;gap:8px;margin-bottom:8px}
.kit-item-name{flex:1;min-width:0;font-size:13.5px;font-weight:700;color:var(--g900)}
.kit-item-ctrl{display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.seg.mini{display:inline-flex;border:1px solid var(--line);border-radius:9px;overflow:hidden}
.seg.mini .seg-b{padding:6px 10px;font-size:12px}
.ent-seg.mini{display:inline-flex;gap:4px}
.ent-seg.mini .ent-b{padding:6px 10px;font-size:12px}
/* ---- Ajouts à une demande (amendments) ---- */
.add-banner{background:#eef6ff;border:1px solid #bcd8f5;border-radius:12px;padding:10px 12px;font-size:12.5px;color:#1b4a7a;margin-bottom:10px;line-height:1.4}
.add-items-btn{width:100%;background:var(--g);color:#fff;border:0;border-radius:12px;padding:12px;font-weight:800;font-size:14px;cursor:pointer;font-family:inherit;margin-bottom:12px}
.prep-add-btn{margin-top:8px;background:#fff;border:1.5px solid var(--g);color:var(--g);border-radius:10px;padding:8px 12px;font-weight:700;font-size:13px;cursor:pointer;font-family:inherit}
.add-tag{display:inline-block;background:#fff4e6;color:#8a5a00;border:1px solid #f0c47a;border-radius:6px;padding:0 5px;font-size:10px;font-weight:700;margin-left:4px;white-space:nowrap;vertical-align:middle}
.doc-adds{margin-top:6px;font-size:10.5px;color:#222;line-height:1.45}
.doc-adds b{color:#8a5a00}
.add-batch u{text-decoration:underline;font-weight:700}
/* ---- Page de notes manuscrites (impression) ---- */
.notes-page{page-break-before:always}
.notes-h{display:flex;align-items:center;gap:8px;font-weight:800;font-size:14px;color:var(--g900);margin-bottom:16px}
.notes-h .hb-logo{height:22px}
.notes-lines{display:flex;flex-direction:column}
.nl{height:32px;border-bottom:1px solid #c9d6cd}
`;
