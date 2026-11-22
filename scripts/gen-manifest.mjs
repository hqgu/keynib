// 生成加载项清单：开发版（localhost:3000）与正式版（GitHub Pages）
// 用法：node scripts/gen-manifest.mjs
// 输出：manifest/manifest.dev.xml、manifest/manifest.prod.xml，以及供下载的 apps/taskpane/public/manifest.xml（= 正式版）
import fs from 'node:fs';

const TARGETS = [
  {
    file: 'manifest/manifest.dev.xml',
    id: 'd6162bce-0713-4aa1-9115-4068427bbde4',
    base: 'https://localhost:3000',
    support: 'https://localhost:3000/taskpane.html',
    nameZh: '妙笔 Keynib (Dev)',
    nameEn: 'Keynib (Dev)',
  },
  {
    file: 'manifest/manifest.prod.xml',
    id: 'ffa53a3c-1586-4ab9-b009-312258005896',
    base: 'https://hqgu.github.io/keynib',
    support: 'https://hqgu.github.io/keynib/',
    nameZh: '妙笔 Keynib',
    nameEn: 'Keynib',
  },
];

const hostBlock = (name) => `        <Host xsi:type="${name}">
          <DesktopFormFactor>
            <GetStarted>
              <Title resid="GetStarted.Title"/>
              <Description resid="GetStarted.Description"/>
              <LearnMoreUrl resid="LearnMore.Url"/>
            </GetStarted>
            <ExtensionPoint xsi:type="PrimaryCommandSurface">
              <OfficeTab id="TabHome">
                <Group id="Keynib.Group">
                  <Label resid="Group.Label"/>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16"/>
                    <bt:Image size="32" resid="Icon.32"/>
                    <bt:Image size="80" resid="Icon.80"/>
                  </Icon>
                  <Control xsi:type="Button" id="Keynib.OpenPane">
                    <Label resid="Button.Label"/>
                    <Supertip>
                      <Title resid="Button.Label"/>
                      <Description resid="Button.Tooltip"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16"/>
                      <bt:Image size="32" resid="Icon.32"/>
                      <bt:Image size="80" resid="Icon.80"/>
                    </Icon>
                    <Action xsi:type="ShowTaskpane">
                      <TaskpaneId>KeynibPane</TaskpaneId>
                      <SourceLocation resid="Taskpane.Url"/>
                    </Action>
                  </Control>
                </Group>
              </OfficeTab>
            </ExtensionPoint>
          </DesktopFormFactor>
        </Host>`;

const manifest = (t) => `<?xml version="1.0" encoding="UTF-8"?>
<!-- 由 scripts/gen-manifest.mjs 生成，请勿手改 -->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
           xsi:type="TaskPaneApp">
  <Id>${t.id}</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Keynib</ProviderName>
  <DefaultLocale>zh-CN</DefaultLocale>
  <DisplayName DefaultValue="${t.nameZh}">
    <Override Locale="en-US" Value="${t.nameEn}"/>
  </DisplayName>
  <Description DefaultValue="自带密钥（BYOK）的 AI 写作、演示与数据分析助手，支持 PowerPoint、Word、Excel。">
    <Override Locale="en-US" Value="Bring-your-own-key AI assistant for PowerPoint, Word and Excel."/>
  </Description>
  <IconUrl DefaultValue="${t.base}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="${t.base}/assets/icon-80.png"/>
  <SupportUrl DefaultValue="${t.support}"/>
  <AppDomains>
    <AppDomain>${new URL(t.base).origin}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Presentation"/>
    <Host Name="Document"/>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="${t.base}/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
${['Presentation', 'Document', 'Workbook'].map(hostBlock).join('\n')}
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="${t.base}/assets/icon-16.png"/>
        <bt:Image id="Icon.32" DefaultValue="${t.base}/assets/icon-32.png"/>
        <bt:Image id="Icon.80" DefaultValue="${t.base}/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="${t.base}/taskpane.html"/>
        <bt:Url id="LearnMore.Url" DefaultValue="${t.support}"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="妙笔 Keynib 已就绪">
          <bt:Override Locale="en-US" Value="Keynib is ready"/>
        </bt:String>
        <bt:String id="Group.Label" DefaultValue="妙笔">
          <bt:Override Locale="en-US" Value="Keynib"/>
        </bt:String>
        <bt:String id="Button.Label" DefaultValue="妙笔助手">
          <bt:Override Locale="en-US" Value="Keynib"/>
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="在“开始”选项卡点击“妙笔助手”打开面板。">
          <bt:Override Locale="en-US" Value="Click Keynib on the Home tab to open the pane."/>
        </bt:String>
        <bt:String id="Button.Tooltip" DefaultValue="打开妙笔 Keynib 面板">
          <bt:Override Locale="en-US" Value="Open the Keynib pane"/>
        </bt:String>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;

for (const t of TARGETS) {
  fs.writeFileSync(t.file, manifest(t));
  console.log('已生成', t.file);
}
fs.copyFileSync('manifest/manifest.prod.xml', 'apps/taskpane/public/manifest.xml');
console.log('已复制正式版清单到 apps/taskpane/public/manifest.xml（供下载）');
