from __future__ import annotations

import threading

from ...auth import *
from ...config import *
from ...api import *
from ...excel import *
from ..dialogs import ConnectionDialog


class ConnectionMixin:
    def _update_connection_summary(self):
        profile = self.profile_combo.currentText() if hasattr(self, "profile_combo") else ""
        base_widget = self.fields.get("BASE_URL") if hasattr(self, "fields") else None
        user_widget = self.fields.get("USERNAME") if hasattr(self, "fields") else None
        base_url = base_widget.text().strip() if base_widget else ""
        username = user_widget.text().strip() if user_widget else ""
        username_text = username if username else "не указан"
        self.connection_summary_label.setText(
            f"Профиль: {profile}   |   BASE_URL: {base_url}   |   USERNAME: {username_text}"
        )

    def _show_connection_dialog(self):
        dialog = ConnectionDialog(self)
        dialog.exec()

    def _toggle_connection_settings(self):
        """Toggle visibility of BASE_URL and advanced connection settings."""
        is_visible = self.connection_details_group.isVisible()
        self.connection_details_group.setVisible(not is_visible)
        if is_visible:
            self.connection_toggle_btn.setText("Показать настройки")
        else:
            self.connection_toggle_btn.setText("Скрыть настройки")

    def _toggle_advanced_settings(self):
        """Toggle visibility of advanced connection settings."""
        is_visible = self.advanced_group.isVisible()
        self.advanced_group.setVisible(not is_visible)
        if is_visible:
            self.advanced_toggle_btn.setText("▼ Расширенные настройки")
        else:
            self.advanced_toggle_btn.setText("▲ Расширенные настройки")

    def _on_profile_changed(self, index):
        """Handle profile selection change."""
        # Get the internal preset key from userData
        profile_key = self.profile_combo.currentData()
        if profile_key not in CONNECTION_PRESETS:
            return

        preset = CONNECTION_PRESETS[profile_key]

        # Update basic fields (only BASE_URL from preset, keep USERNAME/PASSWORD as is)
        self.fields["BASE_URL"].setText(preset["base_url"])
        # Note: USERNAME and PASSWORD are NOT overwritten - user keeps what they typed

        # Update advanced fields
        self.fields["CLIENT_ID"].setText(preset["client_id"])
        self.fields["SSO_BASE_URL"].setText(preset["sso_base_url"])
        self.fields["REALM"].setText(preset["realm"])
        self.fields["BROKER_ALIAS"].setText(preset["broker_alias"])
        self.fields["ADFS_URL"].setText(preset["adfs_base_url"])

        # Show/hide advanced settings based on profile
        is_custom = profile_key == "Custom"
        if is_custom:
            # Show connection details for Custom so BASE_URL and auth params are editable.
            self.connection_details_group.setVisible(True)
            self.connection_toggle_btn.setText("Скрыть настройки")
            self.advanced_toggle_btn.setVisible(True)
            # Keep advanced group visibility as user left it (collapsed or expanded).
        else:
            # Hide advanced settings completely for predefined profiles.
            self.advanced_toggle_btn.setVisible(False)
            self.advanced_group.setVisible(False)
            self.advanced_toggle_btn.setText("▼ Расширенные настройки")

        # For Custom profile, make advanced fields editable; for presets they are read-only
        self.fields["CLIENT_ID"].setReadOnly(not is_custom)
        self.fields["SSO_BASE_URL"].setReadOnly(not is_custom)
        self.fields["REALM"].setReadOnly(not is_custom)
        self.fields["BROKER_ALIAS"].setReadOnly(not is_custom)
        self.fields["ADFS_URL"].setReadOnly(not is_custom)

        # Reset auth status
        self._set_auth_status("Не подключено", "pending")
        self.auth_config_source.setText("")

    def _val(self, key):
        return self.fields[key].text().strip()

    def _get_manual_auth_config(self):
        """Get manually entered auth config from advanced fields."""
        return {
            "client_id": self._val("CLIENT_ID") if self._val("CLIENT_ID") != DEFAULT_CLIENT_ID else None,
            "sso_base_url": self._val("SSO_BASE_URL") if self._val("SSO_BASE_URL") != DEFAULT_SSO_BASE_URL else None,
            "realm": self._val("REALM") if self._val("REALM") != DEFAULT_REALM else None,
            "broker_alias": self._val("BROKER_ALIAS") if self._val("BROKER_ALIAS") != DEFAULT_BROKER_ALIAS else None,
            "adfs_base_url": self._val("ADFS_URL") if self._val("ADFS_URL") != DEFAULT_ADFS_BASE_URL else None,
        }

    def _set_advanced_fields(self, config):
        """Update advanced fields with discovered/cached config."""
        if config.get("client_id"):
            self.fields["CLIENT_ID"].setText(config["client_id"])
        if config.get("sso_base_url"):
            self.fields["SSO_BASE_URL"].setText(config["sso_base_url"])
        if config.get("realm"):
            self.fields["REALM"].setText(config["realm"])
        if config.get("broker_alias"):
            self.fields["BROKER_ALIAS"].setText(config["broker_alias"])
        if config.get("adfs_base_url"):
            self.fields["ADFS_URL"].setText(config["adfs_base_url"])

    def _get_worker_auth_args(self):
        """Get auth args in tuple format expected by workers.
        
        Returns: (auth_mode, client_id, base_url, username, password, sso_url, realm, broker_alias, adfs_url)
        or None if auth config is not available.
        """
        base_url = self._val("BASE_URL")
        username = self._val("USERNAME")
        password = self._val("PASSWORD")
        
        if not base_url or not username or not password:
            return None
        
        # Get current profile key from userData
        current_profile = self.profile_combo.currentData()
        profile_config = CONNECTION_PRESETS.get(current_profile, CONNECTION_PRESETS["Custom"]).copy()
        
        if current_profile == "Custom":
            # For Custom profile, try cache and manual config
            manual_config = self._get_manual_auth_config()
            cached_config = load_connection_profile(base_url)
            config = merge_auth_config(cached_config, None, manual_config)
            
            # If still missing required params, try discovery
            if not all(config.get(k) for k in ["client_id", "sso_base_url", "realm"]):
                discovered = discover_auth_config(base_url)
                config = merge_auth_config(cached_config, discovered, manual_config)
        else:
            # For predefined profiles, use form values if they differ from defaults
            config = profile_config.copy()
            config["client_id"] = self._val("CLIENT_ID") or profile_config["client_id"]
            config["sso_base_url"] = self._val("SSO_BASE_URL") or profile_config["sso_base_url"]
            config["realm"] = self._val("REALM") or profile_config["realm"]
            config["broker_alias"] = self._val("BROKER_ALIAS") or profile_config["broker_alias"]
            config["adfs_base_url"] = self._val("ADFS_URL") or profile_config["adfs_base_url"]
        
        return (
            config.get("auth_mode", "keycloak_broker"),
            config.get("client_id") or DEFAULT_CLIENT_ID,
            base_url,
            username,
            password,
            config.get("sso_base_url") or DEFAULT_SSO_BASE_URL,
            config.get("realm") or DEFAULT_REALM,
            config.get("broker_alias") or DEFAULT_BROKER_ALIAS,
            config.get("adfs_base_url") or DEFAULT_ADFS_BASE_URL,
        )

    def _auth_args_with_discovery(self):
        """Get auth args with auto-discovery support."""
        base_url = self._val("BASE_URL")
        username = self._val("USERNAME")
        password = self._val("PASSWORD")
        
        if not base_url or not username or not password:
            return None
        
        # Get current profile config from userData
        current_profile = self.profile_combo.currentData()
        profile_config = CONNECTION_PRESETS.get(current_profile, CONNECTION_PRESETS["Custom"]).copy()
        
        # For Production and Test presets, use the predefined auth_mode
        # For Custom, try to get manual config
        if current_profile == "Custom":
            manual_config = self._get_manual_auth_config()
            # Try to load from cache
            cached_config = load_connection_profile(base_url)
            # Run auto-discovery if no cache
            discovered_config = None
            if not cached_config or not all(cached_config.get(k) for k in ["client_id", "sso_base_url", "realm"]):
                discovered_config = discover_auth_config(base_url)
            # Merge configs: manual > cached > discovered
            final_config = merge_auth_config(cached_config, discovered_config, manual_config)
            # Use profile's auth_mode if not overridden
            if "auth_mode" not in final_config or not final_config["auth_mode"]:
                final_config["auth_mode"] = profile_config.get("auth_mode", "keycloak_broker")
        else:
            # For predefined profiles, use the profile config directly
            # But allow user to override values in the form fields
            final_config = profile_config.copy()
            final_config["client_id"] = self._val("CLIENT_ID") or profile_config["client_id"]
            final_config["sso_base_url"] = self._val("SSO_BASE_URL") or profile_config["sso_base_url"]
            final_config["realm"] = self._val("REALM") or profile_config["realm"]
            final_config["broker_alias"] = self._val("BROKER_ALIAS") or profile_config["broker_alias"]
            final_config["adfs_base_url"] = self._val("ADFS_URL") or profile_config["adfs_base_url"]
            final_config["source"] = "preset"
        
        return {
            "base_url": base_url,
            "username": username,
            "password": password,
            "config": final_config,
        }

    def _auto_connect(self):
        self._set_auth_status("Автоподключение...", "pending")
        self.auth_config_source.setText("")
        self.auth_btn.setEnabled(False)
        
        auth_data = self._auth_args_with_discovery()
        if not auth_data:
            self._set_auth_status("Заполните BASE_URL, USERNAME и PASSWORD", "danger")
            self.auth_btn.setEnabled(True)
            return
        
        threading.Thread(
            target=self._auth_worker_with_discovery,
            args=(auth_data,),
            daemon=True,
        ).start()

    def _test_auth(self):
        auth_data = self._auth_args_with_discovery()
        if not auth_data:
            self._set_auth_status("Заполните BASE_URL, USERNAME и PASSWORD", "danger")
            return
        
        self.auth_btn.setEnabled(False)
        self._set_auth_status("Проверка...", "pending")
        self.auth_config_source.setText("")
        
        threading.Thread(
            target=self._auth_worker_with_discovery,
            args=(auth_data,),
            daemon=True,
        ).start()

    def _auth_worker_with_discovery(self, auth_data):
        """Auth worker with auto-discovery and caching support."""
        base_url = auth_data["base_url"]
        username = auth_data["username"]
        password = auth_data["password"]
        config = auth_data["config"]
        
        # Get current profile key and display name
        current_profile_key = self.profile_combo.currentData()
        current_profile_display = self.profile_combo.currentText()
        auth_mode = config.get("auth_mode", "keycloak_broker")
        
        try:
            # Log the profile and auth mode
            print(f"Profile: {current_profile_display}")
            print(f"Auth mode: {auth_mode}")
            print(f"Base URL: {base_url}")
            print(f"CLIENT_ID: {config.get('client_id')}")
            
            if auth_mode == "keycloak_broker":
                print(f"SSO_BASE_URL: {config.get('sso_base_url')}")
                print(f"REALM: {config.get('realm')}")
                print(f"BROKER_ALIAS: {config.get('broker_alias')}")
                print(f"ADFS_URL: {config.get('adfs_base_url')}")
            elif auth_mode == "adfs_direct":
                print(f"ADFS_URL: {config.get('adfs_base_url')}")
            
            # Update UI with config info (use display name for user-friendly status)
            self.auth_config_source.setText(f"Профиль: {current_profile_display} | Auth: {auth_mode}")
            
            # Check required parameters based on auth mode
            if auth_mode == "keycloak_broker":
                required_params = ["client_id", "sso_base_url", "realm"]
                missing_params = [p for p in required_params if not config.get(p)]
                
                if missing_params:
                    self.advanced_group.setVisible(True)
                    self.advanced_toggle_btn.setText("▲ Расширенные настройки")
                    self._set_advanced_fields(config)
                    raise Exception(f"Не удалось определить параметры: {', '.join(missing_params)}. Заполните вручную в расширенных настройках.")
                
                # Use Keycloak broker authentication
                authenticate(
                    session,
                    config["client_id"],
                    base_url,
                    username,
                    password,
                    sso_base_url=config.get("sso_base_url"),
                    realm=config.get("realm"),
                    broker_alias=config.get("broker_alias") or DEFAULT_BROKER_ALIAS,
                    adfs_base_url=config.get("adfs_base_url") or DEFAULT_ADFS_BASE_URL,
                )
                
            elif auth_mode == "adfs_direct":
                if not config.get("client_id") or not config.get("adfs_base_url"):
                    raise Exception("Не заданы CLIENT_ID или ADFS_URL для прямой авторизации ADFS.")
                
                # Use direct ADFS OAuth authentication
                authenticate_adfs_direct(
                    session,
                    config["client_id"],
                    base_url,
                    username,
                    password,
                    adfs_base_url=config.get("adfs_base_url"),
                )
            else:
                raise Exception(f"Неизвестный auth_mode: {auth_mode}")
            
            # Save successful config to cache
            save_connection_profile(base_url, config)
            
            self._set_auth_status("Подключено", "success")
            
        except Exception as e:
            error_msg = str(e)
            self._set_auth_status(f"Ошибка: {error_msg}", "danger")
        finally:
            self.auth_btn.setEnabled(True)

    def _set_auth_status(self, text, tone="neutral"):
        self.auth_status.setText(text)
        self.auth_status.setProperty("tone", tone)
        style = self.auth_status.style()
        style.unpolish(self.auth_status)
        style.polish(self.auth_status)
        self.auth_status.update()
        if hasattr(self, "connection_summary_label"):
            self._update_connection_summary()
