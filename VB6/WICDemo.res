        ��  ��                  �      �� ��     0 	        <?xml version="1.0" encoding="UTF-8"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <assemblyIdentity
    version="1.0.0.0"
    processorArchitecture="X86"
    name="WICDemo"
    type="win32"
    />
  <description>WICDemo</description>
    <dependency>
        <dependentAssembly>
            <assemblyIdentity
                type="win32"
                name="Microsoft.Windows.Common-Controls"
                version="6.0.0.0"
                processorArchitecture="X86"
                publicKeyToken="6595b64144ccf1df"
                language="*"
             />
        </dependentAssembly>
    </dependency>
    <dependency>
        <dependentAssembly>
            <assemblyIdentity
                type="win32"
                name="Microsoft.Windows.GdiPlus"
                version="1.1.0.0"
                processorArchitecture="X86"
                publicKeyToken="6595b64144ccf1df"
                language="*"
             />
        </dependentAssembly>
    </dependency>
<!-- Identify the application security requirements: Vista and above -->
  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v2">
      <security>
        <requestedPrivileges>
          <requestedExecutionLevel
            level="asInvoker"
            uiAccess="false"
            />
        </requestedPrivileges>
      </security>
  </trustInfo>
</assembly>