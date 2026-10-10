"""Updater contract checks; Git calls are mocked, never mutate this checkout."""
from pathlib import Path
import subprocess
import sys
import tempfile
from types import SimpleNamespace
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'tools'))
from direct_branch_update import update_checkout,BRANCH


class UpdateTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.addCleanup(self.temp.cleanup)
        self.repo=Path(self.temp.name).resolve();self.commands=[];self.branches=0
        self.changed=False;self.cancelled=False;self.stash_error=False
        self.stashed=False;self.leftovers=False

    def fake_run(self,command,**kwargs):
        self.commands.append(command)
        self.assertEqual(kwargs['cwd'],self.repo)
        args=command[1:];value=''
        if args==['rev-parse','--show-toplevel']:value=str(self.repo)
        elif args==['branch','--show-current']:
            self.branches+=1;value='main' if self.changed and self.branches>1 else BRANCH
        elif args==['rev-parse','HEAD']:value='a'*40
        elif args==['rev-parse','--verify','refs/stash']:value='c'*40
        elif args[0]=='rev-parse':value='b'*40
        elif args==['status','--porcelain']:value=' M local.py\n?? notes.txt' if not self.stashed or self.leftovers else ''
        elif args[:2]==['stash','push']:
            if self.stash_error:raise subprocess.CalledProcessError(1,command)
            self.stashed=True
        return SimpleNamespace(stdout=value)

    def test_backup_before_reset_and_exact_verified_remote_commit(self):
        result=update_checkout(self.repo,run=self.fake_run)
        self.assertEqual(result,dict(commit='b'*40,stash='c'*40))
        stash=next(i for i,c in enumerate(self.commands) if c[1:3]==['stash','push'])
        reset=next(i for i,c in enumerate(self.commands) if c[1]=='reset')
        self.assertLess(stash,reset)
        self.assertEqual(self.commands[reset],['git','reset','--hard','b'*40])
        self.assertFalse(any('clean' in c for c in self.commands))

    def test_changed_branch_and_failed_backup_never_reset(self):
        self.changed=True
        with self.assertRaises(ValueError):update_checkout(self.repo,run=self.fake_run)
        self.assertFalse(any(c[1] in ('stash','reset') for c in self.commands))
        self.changed=False;self.branches=0;self.commands=[];self.stash_error=True
        with self.assertRaises(subprocess.CalledProcessError):update_checkout(self.repo,run=self.fake_run)
        self.assertFalse(any(c[1]=='reset' for c in self.commands))

    def test_cancelled_fetch_does_not_replace_or_stash_checkout(self):
        def run(command,**kwargs):
            result=self.fake_run(command,**kwargs)
            if command[1]=='fetch':self.cancelled=True
            return result
        with self.assertRaises(InterruptedError):update_checkout(self.repo,run=run,cancelled=lambda:self.cancelled)
        self.assertFalse(any(c[1] in ('stash','reset') for c in self.commands))

    def test_wrong_root_never_fetches(self):
        def run(command,**kwargs):return SimpleNamespace(stdout=str(self.repo/'other'))
        with self.assertRaises(ValueError):update_checkout(self.repo,run=run)

    def test_unstashable_changes_abort_with_backup_not_reset(self):
        self.leftovers=True
        with self.assertRaisesRegex(ValueError,'could not be stashed'):update_checkout(self.repo,run=self.fake_run)
        self.assertTrue(self.stashed)
        self.assertFalse(any(c[1]=='reset' for c in self.commands))


if __name__=='__main__':unittest.main()
