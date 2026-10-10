"""Real Qt integration checks, simulation only; skip if PySide6 unavailable."""
import importlib.util
import json
from pathlib import Path
import sys
import tempfile
import threading
import time
import unittest
from unittest.mock import Mock, patch

HAS_QT=importlib.util.find_spec('PySide6') is not None
if HAS_QT:
    sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'tools'))
    from PySide6.QtCore import QTimer
    from PySide6.QtWidgets import QApplication,QPushButton,QCheckBox,QDialog,QComboBox,QMessageBox
    import craniotomy_qt as planner
    from direct_api_controller import StereoDriveController
    from stereodrive_api import Calibration


@unittest.skipUnless(HAS_QT,'Install PySide6 to run actual planner integration tests')
class PlannerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):cls.app=QApplication.instance() or QApplication([])

    def setUp(self):
        self.directory=tempfile.TemporaryDirectory()
        self.root=Path(self.directory.name)
        self.settings_patch=patch.object(planner.CraniotomyWindow,'_config_root_dir',lambda _:self.root)
        self.settings_patch.start()
        self.dialogs_patch=patch.object(planner,'QMessageBox')
        self.dialogs=self.dialogs_patch.start()
        self.window=planner.CraniotomyWindow(StereoDriveController())
        self.window.show()

    def tearDown(self):
        self.window.controller.close()
        self.window.close();self.app.processEvents()
        self.dialogs_patch.stop();self.settings_patch.stop();self.directory.cleanup()

    def connect(self,**kwargs):
        c=self.window.controller
        options=dict(allow_dv=True,allow_piston=True,travel_limits=dict(AP=(-40,40),ML=(-40,40),DV=(-40,40),PISTON=(500,4500)));options.update(kwargs)
        c.connect(Calibration(dict(AP=105280,ML=75864,DV=41767),-8572,dict(AP=0,ML=261,DV=0,PISTON=0)),
                  dict(AP=0,ML=261,DV=0,PISTON=0),self.root/'api.json',**options)
        self.window.refresh_live_position()

    def wait_idle(self):
        deadline=time.monotonic()+4
        while self.window.controller.has_active_motion() and time.monotonic()<deadline:
            self.app.processEvents();time.sleep(.01)
        self.assertFalse(self.window.controller.has_active_motion())

    def test_craniotomy_and_injection_tabs_use_side_by_side_columns(self):
        self.window.resize(1440,900)
        self.window.tabs.setCurrentIndex(1)
        self.window.show();self.app.processEvents()
        self.assertEqual(self.window.craniotomy_load_btn.text(),'Load')
        self.assertEqual(self.window.craniotomy_save_btn.text(),'Save')
        for index in (0,1):
            tab=self.window.tabs.widget(index)
            layout=tab.layout()
            self.assertEqual(layout.count(),2)
            left=layout.itemAt(0).widget()
            right=layout.itemAt(1).widget()
            self.assertGreater(left.width(),0)
            self.assertGreater(right.width(),0)
            self.assertLess(abs(left.width()-right.width()),max(left.width(),right.width())*.35)
        self.assertEqual(self.window.injection_sites_view.parentWidget().title(),'Map')
        self.assertEqual(self.window.inject_up_btn.text(),'Step Syringe Up')
        self.assertEqual(self.window.inject_down_btn.text(),'Step Syringe Down')
        self.assertFalse(hasattr(self.window,'pause_injection_btn'))
        self.assertEqual(self.window.start_injection_btn.text(),'Start Injection Sequence')
        self.assertTrue(self.window.sequence_steps_list.wordWrap())
        self.assertEqual(self.window.direct_limit_edits['PISTON'][0].value(),500)
        self.assertEqual(self.window.direct_limit_edits['PISTON'][1].value(),4500)
        self.assertEqual((self.window.syringe_goto_target_nl.minimum(),self.window.syringe_goto_target_nl.maximum()),(500,4500))
        self.assertEqual(self.window.syringe_goto_btn.text(),'Go To')
        self.assertEqual(self.window.block_test_volume_nl.text(),'20')
        self.assertTrue(self.window.confirm_insertion_check.isChecked())
        checks_panel=self.window.injection_sites_layout.itemAtPosition(3,0).widget()
        self.assertIs(checks_panel.layout().itemAt(0).widget(),self.window.block_check)
        self.assertIs(checks_panel.layout().itemAt(1).widget(),self.window.confirm_insertion_check)
        injection_config=self.window._injection_config_dict()
        self.assertTrue(injection_config['confirm_insertion_enabled'])
        self.window.confirm_insertion_check.setChecked(False)
        self.assertFalse(self.window._injection_config_dict()['confirm_insertion_enabled'])
        self.window.confirm_insertion_check.setChecked(True)
        self.assertNotEqual(self.window.validate_sites_btn.property('variant'),'primary')
        self.assertIs(self.window.injection_sites_layout.itemAtPosition(1,0).widget(),self.window.load_site_set_btn)
        self.assertIs(self.window.injection_sites_layout.itemAtPosition(1,1).widget(),self.window.save_site_set_btn)
        site_action_widths=[button.width() for button in self.window.injection_site_action_buttons]
        site_action_heights={button.height() for button in self.window.injection_site_action_buttons}
        self.assertLessEqual(max(site_action_widths)-min(site_action_widths),1)
        self.assertEqual(len(site_action_heights),1)
        self.assertEqual(self.window.start_injection_btn.width(),self.window.stop_injection_btn.width())
        self.assertEqual(self.window.capture_surface_btn.y(),self.window.move_seed_btn.y())
        self.assertEqual(self.window.capture_surface_btn.x()<self.window.move_seed_btn.x(),True)
        self.assertLess(abs(self.window.capture_surface_btn.width()-self.window.move_seed_btn.width()),4)
        self.assertNotEqual(self.window.capture_surface_btn.property('variant'),'primary')
        self.assertLess(self.window.set_center_btn.x(),self.window.craniotomy_load_btn.x())
        self.assertLess(self.window.craniotomy_load_btn.x(),self.window.craniotomy_save_btn.x())
        self.assertFalse(self.window.drilling_mode_description.isVisible())
        self.assertGreaterEqual(
            self.window.craniotomy_surface_box.geometry().bottom(),
            self.window.craniotomy_surface_box.parentWidget().rect().bottom()-8,
        )

    def test_freeze_modes_lock_craniotomy_edits_and_can_switch_or_exit(self):
        self.window.freeze_draw_btn.click()
        self.assertTrue(self.window.freeze_draw_btn.isChecked())
        self.assertEqual(self.window.freeze_draw_btn.text(),'Inactivate freeze mode')
        self.assertFalse(self.window.mid_ap.isEnabled())
        self.assertFalse(self.window.start_round_btn.isEnabled())
        self.assertTrue(self.window.unfreeze_draw_btn.isEnabled())
        self.assertTrue(self.window.clear_freeze_btn.isEnabled())
        self.assertEqual(self.window.top_view.mode_label,'Freeze mode on')

        self.window.clear_frozen_points()
        self.assertTrue(self.window.freeze_draw_btn.isChecked())
        self.window.unfreeze_draw_btn.click()
        self.assertFalse(self.window.freeze_draw_btn.isChecked())
        self.assertTrue(self.window.unfreeze_draw_btn.isChecked())
        self.assertEqual(self.window.unfreeze_draw_btn.text(),'Inactivate unfreeze mode')
        self.assertTrue(self.window.freeze_draw_btn.isEnabled())
        self.assertTrue(self.window.clear_freeze_btn.isEnabled())
        self.assertEqual(self.window.top_view.mode_label,'Unfreeze mode on')

        self.window.unfreeze_draw_btn.click()
        self.assertFalse(self.window.unfreeze_draw_btn.isChecked())
        self.assertTrue(self.window.mid_ap.isEnabled())
        self.assertTrue(self.window.start_round_btn.isEnabled())
        self.assertEqual(self.window.top_view.mode_label,'')
        self.assertEqual(self.window.freeze_draw_btn.text(),'Freeze Burr Holes')
        self.assertEqual(self.window.unfreeze_draw_btn.text(),'Unfreeze Burr Holes')

    def test_nudge_sites_mode_locks_injection_panel_and_labels_map(self):
        self.window.injection_sites=[planner.InjectionSite(0,0,0)]
        self.window.refresh_injection_sites_list()
        self.window.nudge_all_sites_btn.click()
        self.assertTrue(self.window.nudge_all_sites_active)
        self.assertEqual(self.window.nudge_all_sites_btn.text(),'Disable Nudge of Sites')
        self.assertFalse(self.window.manual_volume_combo.isEnabled())
        self.assertFalse(self.window.start_injection_btn.isEnabled())
        self.assertFalse(self.window.validate_sites_btn.isEnabled())
        self.assertTrue(self.window.nudge_all_sites_btn.isEnabled())
        self.assertEqual(self.window.injection_sites_view.mode_label,'Nudge Sites Mode')
        self.window.nudge_all_sites_btn.click()
        self.assertFalse(self.window.nudge_all_sites_active)
        self.assertEqual(self.window.nudge_all_sites_btn.text(),'Nudge All Sites')
        self.assertTrue(self.window.manual_volume_combo.isEnabled())
        self.assertTrue(self.window.start_injection_btn.isEnabled())
        self.assertEqual(self.window.injection_sites_view.mode_label,'')

    def test_running_injection_locks_setup_but_keeps_progress_pause_and_stop_active(self):
        self.window.set_injection_sequence_controls_active(True)
        for panel in (self.window.manual_control_box,self.window.injection_sites_box,self.window.injection_map_box):
            self.assertFalse(panel.isEnabled())
        for widget in (self.window.single_injection_volume_nl,self.window.insertion_injection_rate_nl_min,
                       self.window.main_injection_rate_nl_min,self.window.block_test_volume_nl,
                       self.window.sequence_steps_list,self.window.injection_save_btn):
            self.assertFalse(widget.isEnabled())
        for widget in (self.window.injection_progress_label,self.window.injection_progress,
                       self.window.injection_site_progress_label,self.window.injection_site_progress,
                       self.window.start_injection_btn,self.window.stop_injection_btn):
            self.assertTrue(widget.isEnabled())
        self.window.set_injection_sequence_controls_active(False)
        self.assertTrue(self.window.manual_control_box.isEnabled())
        self.assertTrue(self.window.injection_sites_box.isEnabled())
        self.assertTrue(self.window.injection_map_box.isEnabled())
        self.assertTrue(self.window.sequence_steps_list.isEnabled())

    def test_running_craniotomy_locks_editing_but_keeps_pause_control(self):
        self.window._set_borehole_controls_locked(True)
        for widget in (
            self.window.mid_ap, self.window.drill_depth, self.window.drilling_mode_combo,
            self.window.generate_seeds_btn, self.window.clear_craniotomy_btn,
            self.window.craniotomy_load_btn, self.window.craniotomy_points_list,
            self.window.set_craniotomy_surface_btn, self.window.zoom_mode_combo,
            self.window.top_view,
        ):
            self.assertFalse(widget.isEnabled(), widget.objectName() or type(widget).__name__)
        self.assertTrue(self.window.start_round_btn.isEnabled())
        self.window._set_borehole_controls_locked(False)
        self.assertTrue(self.window.mid_ap.isEnabled())
        self.assertTrue(self.window.drilling_mode_combo.isEnabled())
        self.assertTrue(self.window.top_view.isEnabled())

    def test_keyboard_nudges_are_ignored_during_injection(self):
        from unittest.mock import Mock
        self.window.injection_thread=Mock()
        self.window.injection_thread.is_alive.return_value=True
        with patch.object(self.window.controller,'nudge_axis') as nudge, \
             patch.object(self.window,'stop_motion') as stop:
            self.window.keyboard_nudge('AP',False,'AP posterior')
        nudge.assert_not_called()
        stop.assert_not_called()

    def test_injection_validation_dialog_wraps_and_orders_equal_buttons(self):
        observed={}
        def inspect_and_close():
            dialog=self.app.activeModalWidget()
            self.assertIsNotNone(dialog)
            observed['width']=dialog.width()
            observed['wrap']=dialog.validation_message_label.wordWrap()
            observed['buttons']=dialog.validation_action_buttons
            observed['sizes']={(button.width(),button.height()) for button in dialog.validation_action_buttons}
            observed['default']=dialog.validation_action_buttons[0].isDefault()
            dialog.validation_action_buttons[-1].click()
        QTimer.singleShot(0,inspect_and_close)
        action,_dialog=self.window._validation_dialog(
            0,1,planner.InjectionSite(ap=0,ml=0,dv=None)
        )
        self.assertEqual(action,'cancel')
        self.assertEqual(observed['width'],max(640,self.window.width()//2))
        self.assertTrue(observed['wrap'])
        self.assertEqual([button.text() for button in observed['buttons']],
                         ['Validate and Next','Next Without Validating','Delete Point','Cancel'])
        self.assertEqual(len(observed['sizes']),1)

    def test_craniotomy_surface_dialog_orders_equal_width_buttons(self):
        observed={}
        def inspect_and_close():
            dialog=self.app.activeModalWidget()
            self.assertIsNotNone(dialog)
            buttons=dialog.findChildren(QPushButton)
            buttons.sort(key=lambda button: button.mapTo(dialog,QPoint(0,0)).x())
            observed['buttons']=buttons
            observed['widths']={button.width() for button in buttons}
            observed['default']=[button.text() for button in buttons if button.isDefault()]
            buttons[-1].click()
        from PySide6.QtCore import QPoint
        QTimer.singleShot(0,inspect_and_close)
        action=self.window._craniotomy_surface_dialog(0,0.0,0.0,-1.0)
        self.assertEqual(action,'cancel')
        self.assertEqual([button.text() for button in observed['buttons']],
                         ['Set and Move to Next','Move to Next','Cancel'])
        self.assertEqual(len(observed['widths']),1)
        self.assertEqual(observed['default'],['Set and Move to Next'])

    def test_start_drilling_prompt_defaults_yes_and_places_it_first(self):
        observed={}
        def inspect_and_choose_yes():
            dialog=self.app.activeModalWidget()
            self.assertIsInstance(dialog,QMessageBox)
            buttons=dialog.buttons()
            observed['labels']=[button.text() for button in buttons]
            observed['default']=dialog.defaultButton().text()
            buttons[0].click()
        with patch.object(planner,'QMessageBox',QMessageBox):
            QTimer.singleShot(0,inspect_and_choose_yes)
            accepted=self.window._confirm_start_drilling_after_surfaces()
        self.assertTrue(accepted)
        self.assertEqual(observed['labels'],['Yes','No'])
        self.assertEqual(observed['default'],'Yes')

    def test_drill_power_prompt_defaults_yes_and_places_it_first(self):
        observed={}
        def inspect_and_choose_yes():
            dialog=self.app.activeModalWidget()
            self.assertIsInstance(dialog,QMessageBox)
            buttons=dialog.buttons()
            observed['labels']=[button.text() for button in buttons]
            observed['default']=dialog.defaultButton().text()
            buttons[0].click()
        with patch.object(planner,'QMessageBox',QMessageBox), \
             patch.object(self.window.controller,'reported_drill_state',side_effect=[False,True]), \
             patch.object(self.window.controller,'prepare_motion'), \
             patch.object(self.window.controller,'set_drill_power'):
            QTimer.singleShot(0,inspect_and_choose_yes)
            turned_on=self.window._ensure_drill_on_for_drilling()
        self.assertTrue(turned_on)
        self.assertEqual(observed['labels'],['Yes','No'])
        self.assertEqual(observed['default'],'Yes')
        self.assertTrue(observed['default'])

    def test_disconnected_keyboard_does_not_move(self):
        with patch.object(self.window,'_focus_is_editable',return_value=False):
            self.window.keyboard_nudge('AP',True,'AP')
        self.assertIsNone(self.window.controller.drive)
        self.dialogs.critical.assert_called()

    def test_motion_options_defaults_apply_persist_and_busy_guard(self):
        for axis, edits in self.window.direct_limit_edits.items():
            self.assertEqual([e.value() for e in edits],[500 if axis=='PISTON' else 0,4500 if axis=='PISTON' else 40])
        self.window.direct_limit_edits['AP'][1].setValue(25)
        self.window.direct_speed_combo.setCurrentText('2')
        self.window._apply_direct_motion_options()
        self.assertEqual(json.loads((self.root/'direct-control.json').read_text())['travel_limits']['AP'],[0,25])
        self.window.close();self.app.processEvents()
        self.window=planner.CraniotomyWindow(StereoDriveController())
        self.assertEqual(self.window.direct_limit_edits['AP'][1].value(),25)
        self.assertEqual(self.window.direct_speed_combo.currentText(),'2')
        self.connect()
        self.window._apply_direct_motion_options()
        self.assertEqual(self.window.controller.drive.travel_limits['AP'],(0,25))
        self.assertEqual(self.window.controller.drive.speed_mm_s,2)
        self.window.direct_speed_combo.setCurrentText('1')
        with patch.object(self.window,'_motion_is_active',return_value=True):
            self.window._apply_direct_motion_options()
        self.assertEqual(self.window.controller.drive.speed_mm_s,2)
        self.window.direct_limit_edits['AP'][0].setValue(30)
        self.window._apply_direct_motion_options()
        self.dialogs.warning.assert_called()
        self.assertEqual(self.window.controller.drive.travel_limits['AP'],(0,25))

    def test_keyboard_absolute_progress_and_internal_bregma(self):
        self.connect()
        self.window.set_local_bregma()
        self.window.keyboard_nudge('AP',True,'AP')
        self.wait_idle()
        self.window.controller.prepare_motion()
        self.assertTrue(self.window._move_through_axis_positions_with_progress([(.02,.02,-.01)],title='test',message='simulated move'))
        self.assertAlmostEqual(self.window.get_bregma_position()[0],.02,delta=1/5225)

    def test_invalid_later_path_leg_is_rejected_before_any_target(self):
        self.connect()
        drive=self.window.controller.drive
        before=len(drive._session.transport.packets)
        with self.assertRaises(ValueError):
            self.window._move_through_axis_positions_with_progress([(0,0,-.1),(40.1,0,-.1)],title='test',message='test')
        self.assertFalse(any(p[1]==0x0c for p in drive._session.transport.packets[before:]))

    def test_protocols_blocked_before_any_hardware_target(self):
        self.connect()
        self.window.start_single_injection();self.window.start_drilling_round();self.window.resume_injection_from_selected()
        self.assertEqual(self.dialogs.warning.call_count,3)
        self.assertFalse(any(p[1]==0x0c for p in self.window.controller.drive._session.transport.packets))

    def test_disconnect_clears_stale_cross_and_readings(self):
        self.connect()
        self.window.refresh_live_position()
        self.window.controller.close()
        self.window.refresh_live_position()
        self.assertEqual(self.window.current_ap_label.text(),'—')
        self.assertIsNone(self.window.top_view.current_point)
        self.assertIsNone(self.window.injection_sites_view.current_point)

    def test_direct_settings_separate_and_survive_startup(self):
        self.window.direct_control_settings=dict(calibration=dict(axis_zero_counts=dict(AP=105280,ML=75864,DV=41767),
            piston_3000_count=-8572,anchor_backlash=dict(AP=0,ML=261,DV=0,PISTON=0)),speed_mm_s=2,
            allow_dv=True,allow_piston=True,allow_drill=False,verified=True,new_reference=True)
        self.window.home_axis=(.1,.2,.3)
        self.window._save_general_settings()
        saved=json.loads((self.root/'direct-control.json').read_text())
        general=json.loads(self.window._general_settings_path().read_text())
        self.assertEqual(saved['speed_mm_s'],2)
        self.assertNotIn('verified',saved);self.assertNotIn('new_reference',saved)
        self.assertNotIn('home_axis',general);self.assertNotIn('axis_zero_reference',general)
        self.window.close();self.app.processEvents()
        self.window=planner.CraniotomyWindow(StereoDriveController())
        self.assertEqual(self.window.direct_control_settings['speed_mm_s'],2)
        self.assertEqual(self.window.home_axis,(.1,.2,.3))
        self.assertIsNone(self.window.controller.drive)

    def test_legacy_direct_fields_migrate_without_loss(self):
        path=self.window._direct_control_settings_path()
        self.window.close();self.app.processEvents()
        path.unlink() # Test-owned temporary file only.
        self.window._general_settings_path().write_text(json.dumps(dict(home_axis=[.1,.2,.3],work_axis=[0,0,0],
            axis_zero_reference=dict(AP=105280,ML=75603,DV=41767))))
        self.window=planner.CraniotomyWindow(StereoDriveController())
        saved=json.loads(path.read_text())
        self.assertEqual(saved['home_axis'],[.1,.2,.3])
        self.assertEqual(saved['axis_zero_reference']['ML'],75603)

    def test_legacy_full_syringe_limits_migrate_to_safe_band(self):
        path=self.window._direct_control_settings_path()
        path.write_text(json.dumps(dict(version=1,live=False,speed_mm_s=1,
            allow_dv=True,allow_piston=True,allow_drill=True,allow_pulsed=True,
            travel_limits=dict(AP=[0,40],ML=[0,40],DV=[0,40],PISTON=[0,5000]))))
        self.window._load_direct_control_settings()
        limits=self.window.direct_control_settings['travel_limits']
        self.assertEqual(limits['PISTON'],(500.0,4500.0))
        self.assertEqual(json.loads(path.read_text())['travel_limits']['PISTON'],[500.0,4500.0])

    def test_wrong_mode_direct_settings_fail_closed_and_are_preserved(self):
        path=self.window._direct_control_settings_path()
        payload=json.loads(path.read_text());payload['live']=True
        original=json.dumps(payload)
        path.write_text(original)
        self.window._load_direct_control_settings()
        self.window._save_general_settings()
        self.assertEqual(self.window.direct_control_settings,{})
        self.assertIsNone(self.window.home_axis)
        self.assertEqual(path.read_text(),original)

    def test_manual_piston_uses_worker_and_verified_position(self):
        self.connect();self.window.set_syringe_position(3000)
        self.window.manual_injection_volume_nl=10
        self.window.manual_syringe_step(True)
        self.assertTrue(self.window.controller.has_active_motion())
        self.wait_idle();self.app.processEvents()
        self.assertAlmostEqual(self.window.current_syringe_position(),3010,delta=1/161.36)

    def test_large_manual_and_test_volumes_are_enabled_and_verified(self):
        self.connect();self.window.set_syringe_position(3000)
        self.window.manual_injection_volume_nl=200
        self.window.manual_syringe_step(True);self.wait_idle();self.app.processEvents()
        self.assertAlmostEqual(self.window.current_syringe_position(),3200,delta=1/161.36)
        self.window.block_test_volume_nl.setText('200')
        self.window.test_for_blockage();self.wait_idle();self.app.processEvents()
        self.assertAlmostEqual(self.window.current_syringe_position(),3000,delta=1/161.36)
        self.assertEqual(self.window._nearest_supported_injection_volume(2000),2000)
        self.assertIn(5,planner.MOVE_SPEED_OPTIONS_MM)

    def test_manual_blockage_test_offers_a_second_test_when_reported_blocked(self):
        self.window.controller.direct_api=False
        self.window.set_syringe_position(3000)
        self.dialogs.question.side_effect=[self.dialogs.Yes,self.dialogs.Yes,self.dialogs.No]
        with patch.object(self.window,'_require_idle',return_value=True), \
             patch.object(self.window,'ensure_syringe_move_allowed'), \
             patch.object(self.window.controller,'syringe_step') as syringe_step, \
             patch.object(self.window,'track_injection_delivery'):
            self.window.test_for_blockage()
        self.assertEqual(syringe_step.call_count,2)
        self.assertEqual(syringe_step.call_args_list[0].args,('20 nl',))
        self.assertEqual(syringe_step.call_args_list[1].kwargs['up'],False)

    def test_empty_fill_modal_uses_real_counter_not_assumed_zero(self):
        self.connect(travel_limits=dict(AP=(0,40),ML=(0,40),DV=(0,40),PISTON=(2985,3035)))
        self.dialogs.warning.return_value=planner.QMessageBox.Cancel
        self.window.empty_syringe()
        self.assertFalse(any(p[1]==0x0c for p in self.window.controller.drive._session.transport.packets))
        self.dialogs.warning.return_value=planner.QMessageBox.Yes
        self.window.empty_syringe()
        self.assertAlmostEqual(self.window.current_syringe_position(),2990,delta=1/161.36)
        self.window.fill_syringe()
        self.assertAlmostEqual(self.window.current_syringe_position(),3030,delta=1/161.36)
        self.assertIsNone(self.window.direct_operation_thread)
        self.assertFalse(self.window.validation_move_active)

    def test_syringe_limit_confirmations_are_short(self):
        from unittest.mock import Mock
        drive=Mock();drive.travel_limits={'PISTON':(500,4500)}
        self.window.controller._require=Mock(return_value=drive)
        self.window.controller.read_injectomate_calibrate_scale_nl=Mock(return_value=2500)
        self.dialogs.question.return_value=planner.QMessageBox.Cancel
        with patch.object(self.window,'_require_idle',return_value=True):
            self.window._direct_syringe_limit(False)
            self.assertEqual(self.dialogs.question.call_args.args[2],"Do you want to empty the syringe?")
            self.window._direct_syringe_limit(True)
            self.assertEqual(self.dialogs.question.call_args.args[2],"Do you want to fill the syringe?")

    def test_direct_operation_can_cancel_before_any_motor_packet(self):
        self.connect();timer=QTimer();timer.setInterval(30)
        def cancel():
            if self.window.validation_move_cancel_callback:
                self.window.validation_move_cancel_callback();timer.stop()
        timer.timeout.connect(cancel);timer.start()
        def operation(cancel,progress):
            for _ in range(20):
                if cancel():return None
                time.sleep(.01)
            return self.window.controller.syringe_step('10 nl',stop_requested=cancel)
        try:self.assertIsNone(self.window._run_direct_operation_with_progress('Cancel Test','Waiting',operation))
        finally:timer.stop()
        self.assertFalse(any(p[1]==0x0c for p in self.window.controller.drive._session.transport.packets))
        self.assertIsNone(self.window.controller.error)

    def test_options_dialog_navigation_does_not_issue_motor_shortcuts(self):
        self.connect();timer=QTimer();timer.setSingleShot(True)
        def navigate():
            from PySide6.QtGui import QKeyEvent
            from PySide6.QtCore import QEvent,Qt
            watched=self.window.options_dialog
            event=QKeyEvent(QEvent.KeyPress,Qt.Key.Key_Up,Qt.NoModifier)
            with patch.object(self.window,'keyboard_nudge') as nudge:
                self.window.eventFilter(watched,event)
                nudge.assert_not_called()
            watched.accept()
        timer.timeout.connect(navigate);timer.start(50)
        self.window.options_dialog.exec()
        self.assertFalse(any(p[1]==0x0c for p in self.window.controller.drive._session.transport.packets))

    def test_validation_modal_retains_step_and_movement_not_syringe_shortcuts(self):
        self.connect();self.window.controller.prepare_motion()
        timer=QTimer();timer.setInterval(50);requested=[False]
        def automate():
            from PySide6.QtGui import QKeyEvent
            from PySide6.QtCore import QEvent,Qt
            dialog=next((d for d in self.window.findChildren(QDialog) if d.isVisible() and d.windowTitle()=='Validate Injection Site'),None)
            if dialog is None:return
            if not requested[0]:
                requested[0]=True
                for binding in ('speed_increase','ap_anterior'):
                    key=self.window.movement_key_bindings[binding]
                    self.window.eventFilter(dialog,QKeyEvent(QEvent.KeyPress,key,Qt.NoModifier))
                self.window.eventFilter(dialog,QKeyEvent(QEvent.KeyPress,Qt.Key_F3,Qt.NoModifier))
            elif not self.window.controller.has_active_motion():
                next(b for b in dialog.findChildren(QPushButton) if b.text()=='Validate and Next').click();timer.stop()
        initial_step=self.window.move_speed_step_mm;timer.timeout.connect(automate);timer.start()
        try:action,dialog=self.window._validation_dialog(0,1,planner.InjectionSite(0,0,None))
        finally:timer.stop()
        self.assertEqual(action,'validate')
        self.assertGreater(self.window.move_speed_step_mm,initial_step)
        self.assertAlmostEqual(self.window.controller.get_current_axis('AP'),self.window.move_speed_step_mm,delta=1/5225)
        self.assertAlmostEqual(self.window.controller.read_injectomate_calibrate_scale_nl(),3000,delta=1/161.36)

    def test_update_disconnects_and_requires_restart_no_real_git_calls(self):
        self.connect();self.dialogs.warning.return_value=planner.QMessageBox.Yes
        self.dialogs.question.return_value=planner.QMessageBox.No
        def update(repo,**kwargs):
            self.assertIsNone(self.window.controller.drive)
            kwargs['progress'](100,'Done')
            return dict(commit='a'*40,stash='b'*40)
        with patch('direct_branch_update.update_checkout',side_effect=update) as updater:
            self.window.update_from_github()
        updater.assert_called_once()
        self.assertTrue(self.window.restart_required)
        self.assertFalse(self.window._require_idle('USB setup'))
        self.assertIn('Git stash',self.window.current_action)

    def test_restart_preserves_live_mode_and_releases_instance_lock(self):
        from unittest.mock import Mock
        lock=Mock();self.window.instance_lock=lock;self.window.controller.live=True
        with patch.object(self.window,'close',return_value=True),patch.object(planner.subprocess,'Popen') as launch:
            self.window._restart_direct_application()
        lock.unlock.assert_called_once()
        self.assertIn('--live',launch.call_args.args[0])
        self.assertTrue(launch.call_args.args[0][1].endswith('craniotomy_qt.py'))
        self.window.controller.live=False

    def test_large_blockage_test_expands_volume_and_tracks_verified_pulses(self):
        self.window.controller.close()
        self.connect(allow_pulsed=True);self.window.controller.prepare_motion()
        self.window.pulsed_clearance_mm=.02
        progress_signal=Mock()
        self.window.injection_progress_signal=progress_signal
        sleeper=time.sleep
        def confirm():self.window.block_prompt_event.set()
        with patch.object(self.window,'block_prompt_signal') as prompt, \
                patch.object(planner.time,'sleep',side_effect=lambda delay:None if delay==1 else sleeper(delay)):
            prompt.emit.side_effect=confirm
            self.window._run_block_test(
                planner.InjectionSite(0,0,0),self.pulse_settings(),200,
                overall_progress_percent=50,
            )
        self.app.processEvents()
        self.assertAlmostEqual(self.window.current_syringe_position(),2800,delta=1/161.36)
        self.assertTrue(progress_signal.emit.call_args_list)
        self.assertTrue(all(call.args[0] == 50 for call in progress_signal.emit.call_args_list))

    def test_successful_retest_repeats_current_site_before_advancing(self):
        self.window.injection_stop_requested.clear()
        first=planner.InjectionSite(0,0,0)
        second=planner.InjectionSite(.1,.1,0)
        first_plan,second_plan=object(),object()
        with (
            patch.object(self.window,'_preflight_pulsed_injections',return_value=[[first_plan],[second_plan]]),
            patch.object(planner.pulsed_protocol,'execute') as execute,
            patch.object(self.window,'_run_block_test',side_effect=[True,False,False]) as blockage,
            patch.object(self.window,'_ensure_repeat_site_capacity') as capacity,
        ):
            self.window._run_pulsed_injections([first,second],self.pulse_settings(),True,20,0,2)
        self.assertEqual([call.args[1] for call in execute.call_args_list],
                         [[first_plan],[first_plan],[second_plan]])
        self.assertEqual(blockage.call_count,3)
        capacity.assert_called_once_with(self.pulse_settings(),20,True)

    def test_block_test_offers_repeat_only_after_a_reported_blockage(self):
        self.window.controller.close()
        self.connect(allow_pulsed=True);self.window.controller.prepare_motion()
        self.window.pulsed_clearance_mm=.02
        prompt_index=[0]
        def answer_block_prompt():
            prompt_index[0]+=1
            self.window.block_prompt_result='retest' if prompt_index[0]==1 else 'clear'
            self.window.block_prompt_event.set()
        def answer_repeat_prompt(site_number):
            self.assertEqual(site_number,2)
            self.window.repeat_injection_site_result=True
            self.window.repeat_injection_site_event.set()
        sleeper=time.sleep
        with (
            patch.object(self.window,'block_prompt_signal') as block_prompt,
            patch.object(self.window,'repeat_injection_site_signal') as repeat_prompt,
            patch.object(planner.time,'sleep',side_effect=lambda delay:None if delay==1 else sleeper(delay)),
        ):
            block_prompt.emit.side_effect=answer_block_prompt
            repeat_prompt.emit.side_effect=answer_repeat_prompt
            repeat=self.window._run_block_test(
                planner.InjectionSite(0,0,0),self.pulse_settings(),20,
                overall_progress_percent=25,site_number=2,
            )
        self.assertTrue(repeat)
        self.assertEqual(prompt_index[0],2)
        repeat_prompt.emit.assert_called_once_with(2)

    def test_repeat_site_prompt_defaults_to_continue(self):
        repeat_button,continue_button=object(),object()
        box=self.dialogs.return_value
        box.addButton.side_effect=[repeat_button,continue_button]
        box.clickedButton.return_value=continue_button
        self.window.repeat_injection_site_event=threading.Event()
        self.window.show_repeat_injection_site_prompt(3)
        self.assertFalse(self.window.repeat_injection_site_result)
        self.assertTrue(self.window.repeat_injection_site_event.is_set())
        self.assertEqual(box.setDefaultButton.call_args.args,(continue_button,))

    def test_repeat_site_prompt_records_repeat_selection(self):
        repeat_button,continue_button=object(),object()
        box=self.dialogs.return_value
        box.addButton.side_effect=[repeat_button,continue_button]
        box.clickedButton.return_value=repeat_button
        self.window.repeat_injection_site_event=threading.Event()
        self.window.show_repeat_injection_site_prompt(1)
        self.assertTrue(self.window.repeat_injection_site_result)
        self.assertTrue(self.window.repeat_injection_site_event.is_set())

    def test_retraction_progress_keeps_the_current_bregma_dv_visible(self):
        self.window.injection_status_site_count=3
        self.window.injection_status_site_index=0
        with patch.object(self.window,'set_status') as status:
            self.window.set_injection_progress(40,'Retracting DV 12.34 mm (Bregma)')
        status.assert_called_once_with('Injection 1/3: Retracting DV 12.34 mm (Bregma)')

    def test_benchmark_options_runs_and_displays_results(self):
        self.connect();timer=QTimer();timer.setInterval(50);started=[False];outputs=[]
        def automate():
            for dialog in self.window.findChildren(QDialog):
                if not dialog.isVisible():continue
                if dialog.windowTitle()=='Supervised Axis Benchmark' and not started[0]:
                    started[0]=True
                    for check in dialog.findChildren(QCheckBox):check.setChecked(check.text()=='AP')
                    from PySide6.QtWidgets import QLineEdit
                    dialog.findChild(QLineEdit).setText('0.01')
                    next(b for b in dialog.findChildren(QPushButton) if b.text()=='Run Benchmark').click()
                elif dialog.windowTitle()=='Axis Movement Benchmark':
                    from PySide6.QtWidgets import QPlainTextEdit
                    outputs.append(dialog.findChild(QPlainTextEdit).toPlainText());dialog.accept();timer.stop()
        timer.timeout.connect(automate);timer.start()
        try:self.window.start_axis_benchmark()
        finally:timer.stop()
        self.assertEqual(len(outputs),1)
        self.assertIn('elapsed_s',outputs[0])
        self.assertEqual(len(outputs[0].strip().splitlines()),3)
        self.assertAlmostEqual(self.window.controller.get_current_axis('AP'),0,delta=1/5225)

    def pulse_settings(self):
        return planner.InjectionProtocolSettings(main_volume_nl=10,insertion_rate_nl_min=60000,
            main_rate_nl_min=60000,injection_depth_mm=.005,insert_retract_speed_um_s=1000,
            overshoot_mm=0,post_inject_pause_s=0)

    def test_pulsed_injection_worker_uses_gui_bregma_and_verified_counter(self):
        self.connect(allow_pulsed=True)
        self.window.controller.prepare_motion();self.window.controller.goto_axis_position(.02,.01,.01)
        self.window.set_local_bregma()
        self.window.validation_clearance_mm=.02
        self.dialogs.warning.reset_mock()
        self.window._start_injection_sequence([planner.InjectionSite(0,0,0)],self.pulse_settings(),[10],False,10,0,1,'test')
        deadline=time.monotonic()+8
        while self.window.injection_thread.is_alive() and time.monotonic()<deadline:
            self.app.processEvents();time.sleep(.01)
        self.assertFalse(self.window.injection_thread.is_alive());self.app.processEvents()
        p=self.window.controller.get_current_axis_position()
        self.assertAlmostEqual(p[0],.02,delta=1/5225)
        self.assertAlmostEqual(p[1],.01,delta=1/5225)
        self.assertAlmostEqual(p[2],-.01,delta=1/5225)
        self.assertAlmostEqual(self.window.current_syringe_position(),2980,delta=1/161.36)
        self.assertIsNone(self.window.controller.error)
        self.dialogs.warning.assert_not_called()

    def test_insertion_confirmation_defaults_yes_and_reverses_extra_travel(self):
        self.window.controller.move_axis_to_target=Mock()
        self.window._wait_for_insertion_confirmation_choice=Mock(side_effect=["increase","confirm"])
        keep_running,_paused=self.window._confirm_pipette_insertion(1.0,self.pulse_settings())
        self.assertTrue(keep_running)
        self.assertEqual(self.window._wait_for_insertion_confirmation_choice.call_count,2)
        self.assertEqual(self.window.controller.move_axis_to_target.call_args_list[0].args[:2],('DV',1.05))
        self.assertEqual(self.window.controller.move_axis_to_target.call_args_list[1].args[:2],('DV',1.0))

    def test_insertion_confirmation_dialog_defaults_to_yes(self):
        confirm_button,increase_button=object(),object()
        box=self.dialogs.return_value
        box.addButton.side_effect=[confirm_button,increase_button,object()]
        box.clickedButton.return_value=confirm_button
        self.window.insertion_confirmation_event=threading.Event()

        self.window.show_insertion_confirmation()

        self.assertEqual(self.window.insertion_confirmation_result,'confirm')
        self.assertTrue(self.window.insertion_confirmation_event.is_set())
        box.setDefaultButton.assert_called_once_with(confirm_button)

    def test_cancel_insertion_confirmation_reverses_extra_travel_and_pauses(self):
        self.window.controller.move_axis_to_target=Mock()
        self.window.start_injection_btn=QPushButton()
        self.window._wait_for_insertion_confirmation_choice=Mock(
            side_effect=["increase","pause","confirm"]
        )
        self.window.injection_pause_requested.clear()
        self.window.injection_stop_requested.clear()

        def resume_after_pause(_seconds):
            self.window.injection_pause_requested.clear()
            self.window.set_injection_paused_ui(False)

        with patch.object(planner.time,'sleep',side_effect=resume_after_pause):
            keep_running,_paused=self.window._confirm_pipette_insertion(1.0,self.pulse_settings())

        self.assertTrue(keep_running)
        self.assertEqual(self.window.controller.move_axis_to_target.call_args_list[0].args[:2],('DV',1.05))
        self.assertEqual(self.window.controller.move_axis_to_target.call_args_list[1].args[:2],('DV',1.0))
        self.assertFalse(self.window.injection_pause_requested.is_set())
        self.assertEqual(self.window.start_injection_btn.text(),'Pause')

    def test_pulsed_workflow_preflight_rejects_later_site_and_whole_dose(self):
        self.connect(allow_pulsed=True);self.window.injection_clearance_axis_dv=-.02
        self.window.pulsed_clearance_mm=.02
        with self.assertRaises(ValueError):
            self.window._preflight_pulsed_injections([planner.InjectionSite(0,0,0),planner.InjectionSite(40.1,0,0)],self.pulse_settings(),False,10)
        with self.assertRaises(Exception):
            self.window._preflight_pulsed_injections([planner.InjectionSite(0,0,0)]*251,self.pulse_settings(),False,10)

        self.window._preflight_pulsed_injections([planner.InjectionSite(0,0,0)]*6,self.pulse_settings(),False,10)
        self.assertFalse(any(p[1]==0x0c for p in self.window.controller.drive._session.transport.packets))

    def test_resume_selected_routes_remaining_sites_to_pulsed_sequence(self):
        self.connect(allow_pulsed=True);self.window.set_local_bregma()
        self.window.injection_sites=[planner.InjectionSite(0,0,0),planner.InjectionSite(.01,.02,0)]
        self.window.refresh_injection_sites_list();self.window.injection_sites_list.setCurrentRow(1)
        with patch.object(self.window,'_injection_protocol_settings',return_value=self.pulse_settings()), \
                patch.object(self.window,'_start_injection_sequence') as start:
            self.window.resume_injection_from_selected()
        self.assertEqual(start.call_args.kwargs['start_site_offset'],1)
        self.assertEqual(len(start.call_args.kwargs['sites']),1)
        self.assertEqual(start.call_args.kwargs['sites'][0].ml,.02)

    def test_pulsed_drill_pause_retracts_without_losing_reference(self):
        self.connect(allow_pulsed=True,allow_drill=True)
        self.window.controller.drive.drill_on()
        self.window.pulsed_clearance_mm=.02;self.window.pulsed_drill_rate=1
        self.window.drill_clearance_axis_dv=-.02
        self.window.drill_pause_requested.set()
        surfaces=[(0,0,0),(.01,0,0),(0,.01,0),(0,0,0)]
        self.window._run_pulsed_drilling(surfaces,[0]*4,[.005]*4,[False]*4,.1,(0,0,-.02))
        self.assertIsNone(self.window.controller.error)
        self.assertTrue(self.window.controller.drive._session.store.read()['valid'])
        self.assertFalse(self.window.controller.drive.drill_state())
        self.assertAlmostEqual(self.window.controller.get_current_axis('DV'),-.02,delta=1/5225)
        self.assertTrue(self.window.drilling_paused)

    def test_pulsed_drilling_start_completes_closed_perimeter(self):
        self.connect(allow_pulsed=True,allow_drill=True)
        self.window.set_local_bregma();self.window.controller.drive.drill_on()
        self.window.validation_clearance_mm=.02
        self.window.trajectory=[(0,0,0),(.01,0,0),(0,.01,0),(0,0,0)]
        self.window.seeds=[planner.SeedPoint(0,0,0,0,0),planner.SeedPoint(1,180,.01,0,0)]
        self.window.drilled_depths=[0]*4;self.window.frozen_points=[False]*4
        self.window.current_target_depth_mm=.005;self.window.drill_depth.setValue(.005)
        self.window.drill_rate_mm_per_s.setValue(1);self.window.round_time_seconds.setValue(1)
        self.window.cut_offset.setValue(0);self.window.mid_ap.setValue(0);self.window.mid_ml.setValue(0)
        self.window.auto_start_rounds.setChecked(False)
        self.dialogs.question.return_value=planner.QMessageBox.Yes
        self.dialogs.warning.return_value=planner.QMessageBox.Yes
        outcomes=[];self.window.drill_round_finished_signal.connect(outcomes.append)
        self.window.start_drilling_round()
        self.assertIsNotNone(self.window.drill_thread)
        deadline=time.monotonic()+10
        while self.window.drill_thread.is_alive() and time.monotonic()<deadline:
            self.app.processEvents();time.sleep(.01)
        self.assertFalse(self.window.drill_thread.is_alive());self.app.processEvents()
        self.assertEqual(outcomes,['completed'])
        self.assertIsNone(self.window.controller.error)
        self.assertEqual(self.window.drilled_depths,[.005]*4)
        self.assertAlmostEqual(self.window.controller.get_current_axis('DV'),-.02,delta=1/5225)

    def test_zero_setup_wizard_simulation(self):
        # Persistent mechanical targets survive reconnect with the same Axis zero.
        self.window.axis_zero_reference=dict(AP=105280,ML=75864-261,DV=41767)
        self.window.home_axis=(.1,0,0)
        invoked=[False]
        timer=QTimer()
        deadline=time.monotonic()+8
        def automate():
            dialog=next((d for d in self.window.findChildren(QDialog) if d.isVisible() and 'measured zero' in d.windowTitle()),None)
            if not dialog:return
            if time.monotonic()>deadline:
                dialog.reject();timer.stop();return
            if not invoked[0]:
                buttons=dialog.findChildren(QPushButton)
                next(b for b in buttons if b.text().startswith('Use SIMULATION')).click()
                next(c for c in dialog.findChildren(QComboBox) if c.count()==2 and c.itemText(0)=='1').setCurrentText('2')
                next(c for c in dialog.findChildren(QCheckBox) if c.text().startswith('Enable DV')).setChecked(True)
                next(c for c in dialog.findChildren(QCheckBox) if c.text().startswith('Measured zero/anchor')).setChecked(True)
                next(b for b in buttons if b.text().startswith('Connect / restore')).click()
                invoked[0]=True
            if self.window.controller.drive is not None:
                button=next(b for b in dialog.findChildren(QPushButton) if b.text()=='Close setup')
                if button.isEnabled():button.click();timer.stop()
        timer.timeout.connect(automate);timer.start(100)
        self.window.open_direct_usb_setup()
        self.assertTrue(invoked[0]);self.assertIsNotNone(self.window.controller.drive)
        self.assertTrue(self.window.controller.drive.position()['calibrated'])
        self.assertEqual(self.window.home_axis,(.1,0,0))
        saved=json.loads(self.window._direct_control_settings_path().read_text())
        self.assertIn('calibration',saved)
        self.assertEqual(saved['speed_mm_s'],2)
        self.assertTrue(saved['allow_dv'])
        self.assertNotIn('verified_backlash',saved)


if __name__=='__main__':unittest.main()
