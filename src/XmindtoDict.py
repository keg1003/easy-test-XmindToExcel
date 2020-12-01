# !/usr/bin/python
# -*- coding: UTF-8 -*-

import xmind
import datetime
import re

def XmindTo(flie_name,user,vban='v1.2.9'):
    workbook = xmind.load(flie_name)
    sheet = workbook.getPrimarySheet()
    root_topic = sheet.getRootTopic()
    sheet1 = root_topic.getData()


    # print(sheet1)
    # print(sheet1["title"])
    # print(len(sheet1["topics"]))
    t_order = []
    t_id = []
    t_point=[]
    t_prec=[]
    t_buzou=[]
    t_resp=[]
    t_name = []

    for i in range(len(sheet1["topics"])):
        title_one=sheet1["topics"]
        if "-" in title_one[i]["title"][8:]:

            new_name=sheet1["topics"][i]["title"][8:].replace("-","，")
        else:
            new_name=sheet1["topics"][i]["title"][8:]
        test_ti=sheet1["topics"][i]["title"][:8]
        # print(test_ti)
        # print(new_name)
        for j in range(len(title_one[i]["topics"])):
            title_two = title_one[i]["topics"]
            test_name1 = title_two[j]["title"]
            cond = title_two[j]["note"]

            for o in range(len(title_two[j]["topics"])):
                title_three = title_two[j]["topics"]
                test_name2 = title_three[o]["title"]
                if "表单" in test_name2:

                    for p in range(len(title_three[o]["topics"])):
                       title_four = title_three[o]["topics"]
                       test_name3 = title_four[p]["title"]



                       # print(title_four[p]["topics"])
                       for u in  range(len(title_four[p]["topics"])):
                           title_five = title_four[p]["topics"]
                           test_point = test_name3+test_name1+test_name2+'-'+title_five[u]["title"]
                           test_name4 = title_five[u]["title"]

                           for y in range(len(title_five[u]["topics"])):
                              title_six = title_five[u]["topics"]
                              respo = title_six[y]['title']
                              # print(respo)
                              respot = respo.split('，')[0]
                              for t in range(len(title_six[y]["topics"])):
                                  title_seven = title_six[y]["topics"]
                                  test_name5 = title_seven[t]['title']
                                  test_buzou = '1.'+title_four[p]["title"]+test_name1+','+title_five[u]["title"]+'输入'+test_name5
                                  test_names = test_name3+test_name1+'，'+test_name4+'输入'+test_name5+'，'+respot
                                  t_buzou.append(test_buzou)
                                  t_order.append(new_name)
                                  t_id.append(test_name1)
                                  t_point.append(test_point)
                                  t_prec.append(cond)
                                  t_resp.append('1.' + respo)
                                  t_name.append(test_names)
                elif "逻辑" in test_name2:
                    test_name2 = title_three[o]["title"]

                    for p in range(len(title_three[o]["topics"])):
                        title_four = title_three[o]["topics"]
                        test_name3 = title_four[p]["title"]


                        # print(title_four[p]["topics"])
                        for u in range(len(title_four[p]["topics"])):
                            title_five = title_four[p]["topics"]
                            test_point = test_name3 + test_name1 + test_name2 + '-' + title_five[u]["title"]
                            respo = title_five[u]["title"]
                            lables = title_five[u]["label"]
                            respot = respo.split('，')[0]
                            if lables:
                                ts = title_five[u]["topics"]
                                lab = [respo,title_five[u]["topics"][0]["title"]]
                                la = []
                                lable = lables.split("；")

                                for l in range(int(lable[0])-2):
                                   ts = ts[0]["topics"]
                                   la.append(ts)

                                   # print(respo,title_five[u]["topics"][0]["title"],title_five[u]["topics"][0]["topics"][0]["title"],title_five[u]["topics"][0]["topics"][0]["topics"][0]["title"])
                                for k in la:
                                       labless=k[0]["title"]
                                       lab.append(labless)
                                buzou = []
                                respo1 = []
                                for w in range(len(lab)):

                                       if w%2==0:
                                           buzou.append(lab[w])
                                       else:

                                           respo1.append(lab[w])
                                buzou1 = []
                                respo2 = []
                                rt = 1
                                for q in buzou:
                                    buzou1.append(str(rt)+'.'+q)
                                    rt += 1
                                rt = 1
                                for q in respo1:
                                    respo2.append(str(rt) + '.' + q)
                                    rt += 1
                                bu = '\n'.join(buzou1)
                                res = '\n'.join(respo2)

                                if '。' in lable[1]:
                                    prepo = '\n'.join(lable[1].split("。"))

                                    prep2 = re.sub("\d\.","","".join(lable[1].split("。")[-1]))

                                else:
                                    prepo = lable[1]
                                test_names = test_name3 + test_name1 + ',' + prep2+ ','+respo1[-1]
                                t_buzou.append(bu)
                                t_order.append(new_name)
                                t_id.append(test_name1)
                                t_point.append(test_point)
                                t_prec.append(prepo)
                                t_resp.append(res)
                                t_name.append(test_names)
                            else:

                                for y in range(len(title_five[u]["topics"])):
                                    title_six = title_five[u]["topics"]
                                    buzou = title_six[y]['title']
                                    if "-" in test_name3:
                                      tit= test_name1 + test_name3
                                    else:
                                      tit = test_name3 + test_name1

                                    test_buzou = '1.' + tit + ',' + buzou
                                    test_names = tit + ',' + buzou+ '，' + respot
                                    t_buzou.append(test_buzou)
                                    t_order.append(new_name)
                                    t_id.append(test_name1)
                                    t_point.append(test_point)
                                    t_prec.append(cond)
                                    t_resp.append('1.'+respo)
                                    t_name.append(test_names)


    t_list = []
    t_path = []
    for e in range(len(t_resp)):
        t_list.append([ t_name[e], '功能用例', vban+'/'+t_id[e], t_prec[e], t_buzou[e], t_resp[e], user, datetime.datetime.now()])
        t_path.append(t_id[e])
    #     print(t_id[e])
    # print(str(datetime.datetime.now()))

    return t_list,t_path

if __name__ == '__main__':
    t = XmindTo('C:/Users/ex-fzk001/Desktop/test.xmind',"fzk","v1.2.9")
    print(t)