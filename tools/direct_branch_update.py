"""Explicit GUI update of the direct branch, with recoverable local changes."""
from pathlib import Path
import subprocess

BRANCH='codex/direct-api-control'


def update_checkout(repo, *, cancelled=lambda:False, progress=lambda percent,text:None,
                    run=subprocess.run):
    repo=Path(repo).resolve()
    def git(*args,timeout=15):
        result=run(['git',*args],cwd=repo,check=True,capture_output=True,text=True,timeout=timeout)
        return result.stdout.strip()
    def checkpoint():
        if cancelled():raise InterruptedError('Update cancelled before replacing the checkout; any local backup remains in Git stash')
    if Path(git('rev-parse','--show-toplevel')).resolve()!=repo:
        raise ValueError('Refusing to update an unexpected repository root')
    if git('branch','--show-current')!=BRANCH:
        raise ValueError(f'Switch to {BRANCH} before using this updater')
    original=git('rev-parse','HEAD')
    checkpoint();progress(10,'Fetching the direct-control branch…')
    git('fetch','origin',f'refs/heads/{BRANCH}:refs/remotes/origin/{BRANCH}',timeout=60)
    target=git('rev-parse','--verify',f'refs/remotes/origin/{BRANCH}^{{commit}}')
    if len(target)!=40 or any(c not in '0123456789abcdef' for c in target):
        raise ValueError('Invalid remote commit identity')
    checkpoint()
    if git('branch','--show-current')!=BRANCH or git('rev-parse','HEAD')!=original:
        raise ValueError('Checkout changed during fetch; update aborted')
    stash=None
    if git('status','--porcelain'):
        progress(60,'Saving local tracked/untracked changes in Git stash…')
        git('stash','push','--include-untracked','-m','Before direct-control GUI update',timeout=30)
        stash=git('rev-parse','--verify','refs/stash')
    checkpoint()
    if git('branch','--show-current')!=BRANCH or git('rev-parse','HEAD')!=original:
        raise ValueError('Checkout changed during backup; update aborted, backup retained')
    if git('status','--porcelain'):
        raise ValueError('Some local changes could not be stashed; update aborted, backup retained')
    progress(85,'Updating the verified direct-control checkout…')
    git('reset','--hard',target)
    progress(100,'Updated. Restart before reconnecting USB.')
    return dict(commit=target,stash=stash)
