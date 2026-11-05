#!/usr/bin/env python3
# Copyright 2026 Google LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

"""
Interactive & CLI Office Add-in XML Manifest Generator
Author: Carlos Augusto, Principal Architect, Google

This script generates a valid, custom Microsoft Office Add-in XML manifest
for Gemini (Agentspace), supporting both Track 1 (WIF) and
Track 2 (Cloud Identity / Google Workspace).
"""

import argparse
import sys
import uuid
import re
import os

def load_env_defaults():
    """Load default values from root or microsoft-addin/.env files."""
    env = {}
    repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    candidates = [
        os.path.join(repo_root, ".env"),
        os.path.join(repo_root, "microsoft-addin", ".env"),
    ]
    for p in candidates:
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        k = k.strip()
                        v = v.strip().strip("'\"")
                        if k and k not in env:
                            env[k] = v
    for k, v in os.environ.items():
        if v and k not in env:
            env[k] = v
    return env

WIF_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>{manifest_id}</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>{provider_name}</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="{display_name}"/>
  <Description DefaultValue="{description}"/>
  <IconUrl DefaultValue="{frontend_url}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="{frontend_url}/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://cloud.google.com/vertex-ai"/>
  <AppDomains>
    <AppDomain>{frontend_url}</AppDomain>
    <AppDomain>{auth_proxy_url}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Document"/>
    <Host Name="Workbook"/>
    <Host Name="Presentation"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="{frontend_url}/taskpane.html?backend=streamassist"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroupWordWIF">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButtonWordWIF">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonIdWordWIF</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>

      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroupExcelWIF">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButtonExcelWIF">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonIdExcelWIF</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>

      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroupPPTWIF">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButtonPPTWIF">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonIdPPTWIF</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="{frontend_url}/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="{frontend_url}/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="{frontend_url}/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://go.microsoft.com/fwlink/?LinkId=276812"/>
        <bt:Url id="Commands.Url" DefaultValue="{frontend_url}/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="{frontend_url}/taskpane.html?backend=streamassist"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="{display_name}"/>
        <bt:String id="CommandsGroup.Label" DefaultValue="{display_name}"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="{button_label}"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="{description}"/>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Open {display_name}."/>
      </bt:LongStrings>
    </Resources>

    <WebApplicationInfo>
      <Id>{client_id}</Id>
      <Resource>api://{frontend_domain}/{client_id}</Resource>
      <Scopes>
        <Scope>profile</Scope>
        <Scope>openid</Scope>
        <Scope>access_as_user</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
</OfficeApp>
"""

GSUITE_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>{manifest_id}</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>{provider_name}</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="{display_name}"/>
  <Description DefaultValue="{description}"/>
  <IconUrl DefaultValue="{frontend_url}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="{frontend_url}/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://cloud.google.com/vertex-ai"/>
  <AppDomains>
    <AppDomain>{frontend_url}</AppDomain>
    <AppDomain>{auth_proxy_url}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Document"/>
    <Host Name="Workbook"/>
    <Host Name="Presentation"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="{frontend_url}/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroupWordGSuite">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButtonWordGSuite">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonIdWordGSuite</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>

      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroupExcelGSuite">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButtonExcelGSuite">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonIdExcelGSuite</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>

      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroupPPTGSuite">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButtonPPTGSuite">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonIdPPTGSuite</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="{frontend_url}/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="{frontend_url}/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="{frontend_url}/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://go.microsoft.com/fwlink/?LinkId=276812"/>
        <bt:Url id="Commands.Url" DefaultValue="{frontend_url}/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="{frontend_url}/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="{display_name}"/>
        <bt:String id="CommandsGroup.Label" DefaultValue="{display_name}"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="{button_label}"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="{description}"/>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Open {display_name}."/>
      </bt:LongStrings>
    </Resources>

    <WebApplicationInfo>
      <Id>{client_id}</Id>
      <Resource>api://{frontend_domain}/{client_id}</Resource>
      <Scopes>
        <Scope>profile</Scope>
        <Scope>openid</Scope>
        <Scope>access_as_user</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
</OfficeApp>
"""

OUTLOOK_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides"
           xsi:type="MailApp">
  <Id>{manifest_id}</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>{provider_name}</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="{display_name}"/>
  <Description DefaultValue="{description}"/>
  <IconUrl DefaultValue="{frontend_url}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="{frontend_url}/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://cloud.google.com/vertex-ai"/>
  <AppDomains>
    <AppDomain>{frontend_url}</AppDomain>
    <AppDomain>{auth_proxy_url}</AppDomain>
    <AppDomain>https://appsforoffice.microsoft.com</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Mailbox"/>
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="{frontend_url}/taskpane.html?host=outlook"/>
        <RequestedHeight>250</RequestedHeight>
      </DesktopSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <SourceLocation DefaultValue="{frontend_url}/taskpane.html?host=outlook"/>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteMailbox</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="ReadOrEdit"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="ReadOrEdit"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Requirements>
        <bt:Sets DefaultMinVersion="1.3">
          <bt:Set Name="Mailbox"/>
        </bt:Sets>
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <DesktopFormFactor>
            <FunctionFile resid="Commands.Url"/>
            <ExtensionPoint xsi:type="MessageReadCommandSurface">
              <OfficeTab id="TabDefault">
                <Group id="msgReadGroup">
                  <Label resid="CommandsGroup.Label"/>
                  <Control xsi:type="Button" id="msgReadOpenPaneButton">
                    <Label resid="TaskpaneButton.Label"/>
                    <Supertip>
                      <Title resid="TaskpaneButton.Label"/>
                      <Description resid="TaskpaneButton.Tooltip"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16x16"/>
                      <bt:Image size="32" resid="Icon.32x32"/>
                      <bt:Image size="80" resid="Icon.80x80"/>
                    </Icon>
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="Taskpane.Url"/>
                      <SupportsPinning>true</SupportsPinning>
                    </Action>
                  </Control>
                </Group>
              </OfficeTab>
            </ExtensionPoint>
            <ExtensionPoint xsi:type="MessageComposeCommandSurface">
              <OfficeTab id="TabDefault">
                <Group id="msgComposeGroup">
                  <Label resid="CommandsGroup.Label"/>
                  <Control xsi:type="Button" id="msgComposeOpenPaneButton">
                    <Label resid="TaskpaneButton.Label"/>
                    <Supertip>
                      <Title resid="TaskpaneButton.Label"/>
                      <Description resid="TaskpaneButton.Tooltip"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16x16"/>
                      <bt:Image size="32" resid="Icon.32x32"/>
                      <bt:Image size="80" resid="Icon.80x80"/>
                    </Icon>
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="Taskpane.Url"/>
                      <SupportsPinning>true</SupportsPinning>
                    </Action>
                  </Control>
                </Group>
              </OfficeTab>
            </ExtensionPoint>
            <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
              <OfficeTab id="TabDefault">
                <Group id="apptReadGroup">
                  <Label resid="CommandsGroup.Label"/>
                  <Control xsi:type="Button" id="apptReadOpenPaneButton">
                    <Label resid="TaskpaneButton.Label"/>
                    <Supertip>
                      <Title resid="TaskpaneButton.Label"/>
                      <Description resid="TaskpaneButton.Tooltip"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16x16"/>
                      <bt:Image size="32" resid="Icon.32x32"/>
                      <bt:Image size="80" resid="Icon.80x80"/>
                    </Icon>
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="Taskpane.Url"/>
                      <SupportsPinning>true</SupportsPinning>
                    </Action>
                  </Control>
                </Group>
              </OfficeTab>
            </ExtensionPoint>
            <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
              <OfficeTab id="TabDefault">
                <Group id="apptComposeGroup">
                  <Label resid="CommandsGroup.Label"/>
                  <Control xsi:type="Button" id="apptComposeOpenPaneButton">
                    <Label resid="TaskpaneButton.Label"/>
                    <Supertip>
                      <Title resid="TaskpaneButton.Label"/>
                      <Description resid="TaskpaneButton.Tooltip"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16x16"/>
                      <bt:Image size="32" resid="Icon.32x32"/>
                      <bt:Image size="80" resid="Icon.80x80"/>
                    </Icon>
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="Taskpane.Url"/>
                      <SupportsPinning>true</SupportsPinning>
                    </Action>
                  </Control>
                </Group>
              </OfficeTab>
            </ExtensionPoint>
          </DesktopFormFactor>
        </Host>
      </Hosts>
      <Resources>
        <bt:Images>
          <bt:Image id="Icon.16x16" DefaultValue="{frontend_url}/assets/icon-16.png"/>
          <bt:Image id="Icon.32x32" DefaultValue="{frontend_url}/assets/icon-32.png"/>
          <bt:Image id="Icon.80x80" DefaultValue="{frontend_url}/assets/icon-80.png"/>
        </bt:Images>
        <bt:Urls>
          <bt:Url id="Commands.Url" DefaultValue="{frontend_url}/commands.html"/>
          <bt:Url id="Taskpane.Url" DefaultValue="{frontend_url}/taskpane.html?host=outlook"/>
        </bt:Urls>
        <bt:ShortStrings>
          <bt:String id="CommandsGroup.Label" DefaultValue="{display_name}"/>
          <bt:String id="TaskpaneButton.Label" DefaultValue="{button_label}"/>
        </bt:ShortStrings>
        <bt:LongStrings>
          <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Open {display_name} for Outlook."/>
        </bt:LongStrings>
      </Resources>
      <WebApplicationInfo>
        <Id>{client_id}</Id>
        <Resource>api://{frontend_domain}/{client_id}</Resource>
        <Scopes>
          <Scope>profile</Scope>
          <Scope>openid</Scope>
          <Scope>access_as_user</Scope>
        </Scopes>
      </WebApplicationInfo>
    </VersionOverrides>
  </VersionOverrides>
</OfficeApp>
"""

def clean_url(url: str) -> str:
    url = url.strip().rstrip('/')
    if not url.startswith('http://') and not url.startswith('https://'):
        url = 'https://' + url
    return url

def extract_domain(url: str) -> str:
    cleaned = clean_url(url)
    return cleaned.replace('https://', '').replace('http://', '').split('/')[0]

def derive_outlook_path(out_file: str) -> str:
    if out_file.endswith(".xml"):
        base = out_file[:-4]
        if base.endswith("-outlook"):
            return out_file
        return f"{base}-outlook.xml"
    return f"{out_file}-outlook.xml"

def prompt_interactive():
    print("=" * 70)
    print("  🚀 Microsoft Office 365 Add-in Manifest Generator")
    print("  Author: Carlos Augusto, Principal Architect, Google")
    print("=" * 70)
    print()

    # 1. Track selection
    print("Choose your Deployment Track:")
    print("  1) Track 1: Workforce Identity Federation (WIF) - Seamless SSO")
    print("  2) Track 2: Cloud Identity / Google Workspace - 3-Legged OAuth")
    track_choice = input("Select [1 or 2] (default: 1): ").strip()
    track = "wif" if track_choice in ("", "1") else "gsuite"

    env = load_env_defaults()
    default_frontend = env.get("FRONTEND_URL") or (f"https://gemini-frontend-{env.get('PROJECT_ID')}.{env.get('REGION', 'us-central1')}.run.app" if env.get("PROJECT_ID") else "")
    default_auth_proxy = env.get("GEMINI_PROXY_URL", "").replace("/askGeminiEnterprise", "") or (f"https://auth-proxy-{env.get('PROJECT_ID')}.{env.get('REGION', 'us-central1')}.run.app" if env.get("PROJECT_ID") else default_frontend)
    default_client_id = env.get("MICROSOFT_ENTRA_APP_ID", "")

    # 2. Frontend URL
    while True:
        prompt_txt = f"\nEnter your Cloud Run gemini-frontend URL (default: '{default_frontend}'): " if default_frontend else "\nEnter your Cloud Run gemini-frontend URL (e.g. https://gemini-frontend-xxx.run.app): "
        frontend_input = input(prompt_txt).strip()
        frontend_val = frontend_input or default_frontend
        if frontend_val:
            frontend_url = clean_url(frontend_val)
            break
        print("❌ Error: gemini-frontend URL is required.")

    # 3. Auth Proxy URL (Optional, defaults to deriving or frontend domain)
    prompt_ap = f"Enter your Cloud Run auth-proxy URL (default: '{default_auth_proxy}'): " if default_auth_proxy else "Enter your Cloud Run auth-proxy URL (optional, press Enter to skip): "
    auth_proxy_input = input(prompt_ap).strip()
    auth_proxy_val = auth_proxy_input or default_auth_proxy or frontend_url
    auth_proxy_url = clean_url(auth_proxy_val)

    # 4. Microsoft Entra Client ID
    while True:
        prompt_cid = f"\nEnter your Microsoft Entra Application (Client) ID (default: '{default_client_id}'): " if default_client_id else "\nEnter your Microsoft Entra Application (Client) ID: "
        client_input = input(prompt_cid).strip()
        client_id = client_input or default_client_id
        if client_id:
            break
        print("❌ Error: Entra Client ID is required.")

    # 5. Display Name
    default_name = "Gemini (WIF)" if track == "wif" else "Gemini"
    display_name_input = input(f"\nAdd-in Display Name (default: '{default_name}'): ").strip()
    display_name = display_name_input if display_name_input else default_name

    # 6. Button Label
    default_button = "Gemini (WIF)" if track == "wif" else "Gemini"
    button_input = input(f"Ribbon Button Label (default: '{default_button}'): ").strip()
    button_label = button_input if button_input else default_button

    # 7. Provider Name
    provider_input = input("Provider / Organization Name (default: 'Google Cloud'): ").strip()
    provider_name = provider_input if provider_input else "Google Cloud"

    # 8. Output filename
    default_out = "microsoft-addin/manifest-gemini.xml"
    out_input = input(f"\nOutput filename (default: '{default_out}'): ").strip()
    out_file = out_input if out_input else default_out

    office_path, manifest_id = generate_manifest(
        track=track,
        frontend_url=frontend_url,
        auth_proxy_url=auth_proxy_url,
        client_id=client_id,
        display_name=display_name,
        button_label=button_label,
        provider_name=provider_name,
        out_file=out_file,
        host="office"
    )
    outlook_out = derive_outlook_path(out_file)
    outlook_path, _ = generate_manifest(
        track=track,
        frontend_url=frontend_url,
        auth_proxy_url=auth_proxy_url,
        client_id=client_id,
        display_name=display_name,
        button_label=button_label,
        provider_name=provider_name,
        out_file=outlook_out,
        host="outlook",
        manifest_id=manifest_id
    )
    return office_path

def generate_manifest(track, frontend_url, auth_proxy_url, client_id, display_name, button_label, provider_name, out_file, host="office", manifest_id=None):
    frontend_url = clean_url(frontend_url)
    auth_proxy_url = clean_url(auth_proxy_url)
    frontend_domain = extract_domain(frontend_url)
    
    # Preserve existing Manifest ID if file already exists to prevent duplicate add-ins in Office
    if not manifest_id:
        candidate_files = [out_file]

        for cand in candidate_files:
            if os.path.exists(cand):
                try:
                    with open(cand, "r", encoding="utf-8") as f:
                        m = re.search(r"<Id>([a-fA-F0-9-]+)</Id>", f.read())
                        if m:
                            manifest_id = m.group(1)
                            break
                except Exception:
                    pass
    if not manifest_id:
        manifest_id = str(uuid.uuid4())
    
    if host == "outlook":
        description = f"{display_name} for Microsoft Outlook."
        template = OUTLOOK_TEMPLATE
    else:
        description = f"{display_name} for Microsoft Word, PowerPoint, and Excel."
        template = WIF_TEMPLATE if track == "wif" else GSUITE_TEMPLATE

    content = template.format(
        manifest_id=manifest_id,
        provider_name=provider_name,
        display_name=display_name,
        button_label=button_label,
        description=description,
        frontend_url=frontend_url,
        frontend_domain=frontend_domain,
        auth_proxy_url=auth_proxy_url,
        client_id=client_id
    )

    with open(out_file, "w", encoding="utf-8") as f:
        f.write(content.strip() + "\n")

    print()
    print("=" * 70)
    print(f"  ✅ Manifest successfully created: {out_file}")
    print("=" * 70)
    print(f"  • Unique Manifest GUID: {manifest_id}")
    print(f"  • Entra App Client ID:  {client_id}")
    print(f"  • Application ID URI:   api://{frontend_domain}/{client_id}")
    print(f"  • Taskpane Source:      {frontend_url}/taskpane.html")
    print("=" * 70)
    print("Next step: Sideload or deploy this manifest in the Microsoft 365 Admin Center.")
    return out_file, manifest_id

def main():
    parser = argparse.ArgumentParser(description="Generate Microsoft Office Add-in XML Manifest for Gemini.")
    parser.add_argument("--track", choices=["wif", "gsuite"], help="Deployment track: 'wif' or 'gsuite'")
    parser.add_argument("--dev", action="store_true", help="Generate a local development manifest pointing to localhost:3000")
    parser.add_argument("--host", choices=["all", "office", "outlook"], default="all", help="Target host family: 'all' (generates both Office & Outlook), 'office', or 'outlook'")
    parser.add_argument("--frontend-url", help="Public HTTPS URL of gemini-frontend Cloud Run service")
    parser.add_argument("--auth-proxy-url", help="Public HTTPS URL of auth-proxy Cloud Run service")
    parser.add_argument("--client-id", help="Microsoft Entra Application (Client) ID")
    parser.add_argument("--display-name", help="Add-in Display Name on Ribbon")
    parser.add_argument("--button-label", help="Ribbon Button Label")
    parser.add_argument("--provider-name", default="Google Cloud", help="Provider/Organization Name")
    parser.add_argument("--manifest-id", help="Explicit Manifest GUID (if not provided, preserves existing file ID or generates fresh UUID)")
    parser.add_argument("--output", help="Output file path (e.g. manifest-custom.xml)")
    parser.add_argument("--interactive", action="store_true", help="Force interactive prompt mode")

    args = parser.parse_args()
    env = load_env_defaults()

    if args.interactive:
        prompt_interactive()
        return

    # Auto-load defaults from .env if not supplied via command line
    frontend_url = args.frontend_url or env.get("FRONTEND_URL")
    if not frontend_url and env.get("PROJECT_ID"):
        region = env.get("REGION", "us-central1")
        frontend_url = f"https://gemini-frontend-{env['PROJECT_ID']}.{region}.run.app"

    client_id = args.client_id or env.get("MICROSOFT_ENTRA_APP_ID")

    auth_proxy_url = args.auth_proxy_url or env.get("GEMINI_PROXY_URL", "").replace("/askGeminiEnterprise", "")
    if not auth_proxy_url and env.get("PROJECT_ID"):
        region = env.get("REGION", "us-central1")
        auth_proxy_url = f"https://auth-proxy-{env['PROJECT_ID']}.{region}.run.app"

    # If neither CLI nor .env provides required params and no args passed, prompt interactively
    if len(sys.argv) == 1 and (not frontend_url or not client_id):
        prompt_interactive()
        return

    track = args.track or ("gsuite" if env.get("GOOGLE_OAUTH_CLIENT_ID") else "wif")
    if args.dev:
        frontend_url = args.frontend_url or "https://localhost:3000"
        auth_proxy_url = args.auth_proxy_url or "https://localhost:3000"
        display_name = args.display_name or "Gemini (Local Debug)"
        button_label = args.button_label or "Gemini (Local Debug)"
        out_file = args.output or "manifest-dev.xml"
        track = args.track or "gsuite"
        client_id = args.client_id or env.get("MICROSOFT_ENTRA_APP_ID") or "00000000-0000-0000-0000-000000000000"
    else:
        display_name = args.display_name or "Gemini"
        button_label = args.button_label or "Gemini"
        auth_proxy_url = auth_proxy_url or frontend_url
        out_file = args.output or "manifest-gemini.xml"

    if not frontend_url or not client_id:
        print("❌ Error: --frontend-url and --client-id could not be resolved from CLI arguments or .env file.", file=sys.stderr)
        print("Please specify them as arguments, run with --interactive, or define them in .env:", file=sys.stderr)
        print("  FRONTEND_URL (or PROJECT_ID) and MICROSOFT_ENTRA_APP_ID", file=sys.stderr)
        sys.exit(1) 

    print("📄 Loaded configuration parameters from .env and arguments:")
    print(f"  • Frontend URL:    {frontend_url}")
    print(f"  • Auth Proxy URL:  {auth_proxy_url}")
    print(f"  • Entra Client ID: {client_id}")
    print(f"  • Track:           {track}")
    print(f"  • Output file:     {out_file}\n")

    target_host = getattr(args, "host", "all")
    initial_manifest_id = args.manifest_id or env.get("OFFICE_ADDIN_MANIFEST_ID")

    if target_host in ("all", "office"):
        office_path, resolved_manifest_id = generate_manifest(
            track=track,
            frontend_url=frontend_url,
            auth_proxy_url=auth_proxy_url,
            client_id=client_id,
            display_name=display_name,
            button_label=button_label,
            provider_name=args.provider_name,
            out_file=out_file,
            host="office",
            manifest_id=initial_manifest_id
        )
    else:
        office_path, resolved_manifest_id = None, initial_manifest_id

    if target_host in ("all", "outlook"):
        outlook_out = out_file if target_host == "outlook" else derive_outlook_path(out_file)
        generate_manifest(
            track=track,
            frontend_url=frontend_url,
            auth_proxy_url=auth_proxy_url,
            client_id=client_id,
            display_name=display_name,
            button_label=button_label,
            provider_name=args.provider_name,
            out_file=outlook_out,
            host="outlook",
            manifest_id=resolved_manifest_id
        )

if __name__ == "__main__":
    main()
