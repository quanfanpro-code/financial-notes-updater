# coding: utf-8
"""正常启动文件的无界面入口，与桌面窗口共用五步处理代码。"""
import argparse
from collections import Counter
from datetime import datetime
import json
from pathlib import Path
import threading
import uuid

from .分步流程 import run_step
from .流程 import save_json
from .表格 import file_hash
from .设置 import DEFAULT_SETTINGS_PATH, load_settings, validate_settings

ORDER=('extract_a','recognize_a','recognize_b','match','write')
SETTINGS=('word_path','b_path','gold_path','base_url','model','timeout','thinking_mode',
          'model_reasoning_effort','report_scope','a_context_text','b_context_text','freeze_standard',
          'update_purpose','comparison_requirements')


def run_background(config, log=None, cancel=None, *, resume=None):
    if not isinstance(config.get('freeze_standard',False),bool):raise ValueError('冻结标准设置须为布尔值')
    settings=validate_settings({k:config[k] for k in (*SETTINGS,'api_key') if k in config})
    inputs={}
    for key in ('word_path','b_path','gold_path'):
        path=Path(settings.get(key) or '').resolve()
        if not path.is_file():raise ValueError('无界面作业缺少原始输入：'+key)
        settings[key]=str(path);inputs[str(path)]=file_hash(path)
    public={k:settings[k] for k in SETTINGS if k in settings}
    done=set();files={};usage=Counter();status='complete'
    if resume:
        folder=Path(resume).resolve()
        manifest=json.loads((folder/'输入记录.json').read_text(encoding='utf-8-sig'))
        if manifest.get('schema')!='附注无界面作业-v1' or manifest['inputs']!=inputs or manifest['settings']!=public:
            raise ValueError('原件或输入设置已变化，不能接续此前作业')
        for journal in sorted(folder.glob('运行事件_*.jsonl')):
            for line in journal.read_text(encoding='utf-8').splitlines():
                row=json.loads(line)
                if row['event'] in ('progress','result'):
                    files.update(row['result'].get('files',{}))
                if row['event']=='result':
                    usage.update(row['result'].get('usage',{}))
                    if row['result']['status'] in ('complete','partial'):
                        done.add(row['step'])
                        if row['result']['status']=='partial':status='partial'
        # 明确选入的新合集接替旧结果，下游重新搬运；原作业记录不改写。
        for carrier in ('a','b'):
            mapping=carrier+'_mapping';answers=carrier+'_scope_answers'
            if config.get(mapping) and config[mapping]!=files.get(mapping):
                done.difference_update(ORDER[ORDER.index('recognize_'+carrier):])
            if 'recognize_'+carrier in done:continue
            if config.get(mapping):files[mapping]=config[mapping]
            if config.get(answers):settings[answers]=config[answers]
    else:
        folder=Path(config.get('output_dir') or Path.cwd()/'输出').resolve()/('后台作业_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'_'+uuid.uuid4().hex[:6])
        folder.mkdir(parents=True)
        save_json(folder/'输入记录.json',{'schema':'附注无界面作业-v1','inputs':inputs,'settings':public,
            'entry':'启动程序.pyw --无界面','initial_state':'仅原始Word、原始Excel及通用标准；不载入历史映射、旧清单或人工安排'})
    settings.update(output_dir=str(folder))
    cancel=cancel or threading.Event();last={}
    stamp=datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    with (folder/('运行事件_'+stamp+'.jsonl')).open('x',encoding='utf-8') as journal:
        def event(kind,**values):
            journal.write(json.dumps({'time':datetime.now().isoformat(),'event':kind,**values},ensure_ascii=False)+'\n');journal.flush()
        def report(message):
            message=str(message)
            if settings.get('api_key'):message=message.replace(settings['api_key'],'[密钥已隐藏]')
            event('log',message=message)
            if log:log(message)
        event('start',resume=bool(resume),output_dir=str(folder))
        try:
            for step in ORDER:
                if step in done:continue
                for path,digest in inputs.items():
                    if file_hash(path)!=digest:raise ValueError('作业原件已变化：'+path)
                settings.update({k:v for k,v in files.items() if k in {'a_path','a_link','a_mapping','b_mapping','update_plan','gold_path'}})
                report('开始：'+step)
                event('step',step=step)
                def progress(result):
                    files.update(result.get('files',{}));event('progress',step=step,result=result)
                last=run_step(step,settings,report,cancel,progress)
                files.update(last.get('files',{}));usage.update(last.get('usage',{}))
                event('result',step=step,result=last)
                # 部分成果继续交给后续步骤使用；真实未决仍随成果保留。
                if last['status']=='partial':status='partial'
                elif last['status']!='complete':
                    status=last['status'];break
            for path,digest in inputs.items():
                if file_hash(path)!=digest:raise ValueError('作业原件已变化：'+path)
        except Exception as error:
            status='failed';last={'message':str(error)}
            report(str(error))
        result={'status':status,'message':last.get('message','本作业此前已处理结束'),'output_dir':str(folder),
            'files':files,'usage':dict(usage),'usage_scope':'步骤返回用量；失败调用原始用量另保存在Codex事件记录中',
            'acceptance':'运行状态只记录步骤是否执行和保存；未明确内容随成果保留，程序不另行裁定语义或来源。'}
        path=save_json(folder/('运行结果_'+stamp+'.json'),result)
        event('finish',result_path=path,status=status)
        report('后台作业结果：'+status+'；'+path)
        return result


def main(argv=None):
    parser=argparse.ArgumentParser(description='从正常设置中的原始Word和原始Excel执行全流程，不操作鼠标键盘。')
    parser.add_argument('--无界面',action='store_true')
    parser.add_argument('--配置',default=str(DEFAULT_SETTINGS_PATH))
    parser.add_argument('--继续',dest='resume')
    parser.add_argument('--冻结标准',action='store_true',help='保持本次标准定义，语义未决只记录，不增补或修订标准。')
    args=parser.parse_args(argv)
    import sys
    def log(message):
        if sys.stdout:print(message,flush=True)
    try:
        settings=load_settings(args.配置)
        if args.冻结标准:settings['freeze_standard']=True
        result=run_background(settings,log=log,resume=args.resume)
    except Exception as error:
        log('后台作业未启动：'+str(error));return 1
    return 0 if result['status']=='complete' else 2
