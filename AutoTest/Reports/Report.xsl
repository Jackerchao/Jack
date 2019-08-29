<?xml version="1.0" ?>
<xsl:stylesheet  version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
    <xsl:template match="/">
        <HTML>
			<HEAD>
                <LINK REL="StyleSheet" HREF="Report.css" TYPE="text/css"/>
                <SCRIPT language = "javascript" type="text/javascript">
                    //Toggle Vertical Menu
					function toggleMenu(id,b)
                    {
                        if (document.getElementById)
                        {
                            var e = document.getElementById(id);
                            var b = document.getElementById(b);
                        
                            if (e)
                            {
                                if (e.style.display != "block")
                                {
                                    e.style.display = "block";
                                    b.src = '_images/check-min.jpg';
                                }
								else
								{
									e.style.display = "none";
                                    b.src = '_images/check-plus.jpg';
								}
                            }
						}
                    }
					function expandall()     
                    {
                        var e = document.all.tags("DIV");
                        var b = document.all.tags("img");
                        for (var i = 0; i &lt; e.length; i++)
                        {
                            if (e[i].style.display == "none")
                            {
                                e[i].style.display = "block";
                            }
                        }
                        for (var i = 0; i &lt; b.length; i++)
                        {
                            if (b[i].id != "m1" &amp;&amp; b[i].id != "m2")
                            {
                                if (b[i].src.substring(b[i].src.lastIndexOf("_"), b[i].src.length) == "_images/check-plus.jpg")
                                {
                                    b[i].src = "_images/check-min.jpg";
                                }
                            }
                        }
                    }
					function collapseall()      
                    {
                        var e = document.all.tags("DIV");
                        var b = document.all.tags("img");

                        for (var i = 0; i &lt; e.length; i++)
                        {
                            if (e[i].id != "Batch" &amp;&amp; e[i].id != "Test" &amp;&amp; e[i].id != "BP" &amp;&amp; e[i].id != "Result")
                            {
                                if (e[i].style.display == "block")
                                {
                                    e[i].style.display = "none";
                                }
                            }
                        }
                        for (var i = 0; i &lt; b.length; i++)
                        {
                            if (b[i].id != "m1" &amp;&amp; b[i].id != "m2" )
                            {
                                if (b[i].src.substring(b[i].src.lastIndexOf("_"), b[i].src.length) == "_images/check-min.jpg")
                                {
                                    b[i].src = "_images/check-plus.jpg";
                                }
                            }
                        }
                    }

                </SCRIPT>
            </HEAD>
            <BODY>
                <table bgcolor="#6699FF" bordercolorlight="#COCOCO" brodercolordark="#COCOCO" border="1" width="100%" height="53">
                    <tr>
                        <td valign="top" height="51" width="1700">
                            <p align="center">
                                <font color="#FFFFFF" face="Garamond" style="font-size: 30pt">
                                    <strong>
                                        Test Automation Accelerator
                                    </strong>
                                </font>
                            </p>
                        </td>
                    </tr>
                </table>
                <table width="100%" height="49" ID="Table2">
                    <tr>
                        <td width="80%" height="43">
                            <p align="Center">
                                <font face="Arial">
                                    <strong>
                                        <big>
                                            <big>Test Automation Report</big>
                                        </big>
                                    </strong>
                                </font>
                            </p>
                        </td>
                    </tr>
                </table>
                <BR></BR>
                <xsl:apply-templates select="Report/TestSuite"></xsl:apply-templates>
            </BODY>
			
        </HTML>
    </xsl:template>
    <xsl:template match="Report/TestSuite">
        <Table id="tblTimeStamp">
            <tr>
                <td>
                    <p align="Left">
                        <font face="Times New Roman" color="blue">
                            <strong>Start Time :</strong>
                        </font>
                    </p>
                </td>
                <td>
                    <xsl:text> </xsl:text>
                    <xsl:value-of select="@StartTime"></xsl:value-of>
                </td>
            </tr>
            <tr>
                <td>
                    <p align="Left">
                        <font face="Times New Roman" color="blue">
                            <strong>End Time :</strong>
                        </font>
                    </p>
                </td>
                <td>
                    <xsl:text></xsl:text>
                    <xsl:value-of select="@EndTime"></xsl:value-of>
                </td>
            </tr>
        </Table>
        <BR></BR>
        <Table id="TableSummary" cellSpacing="1" cellPadding="1" Class="TABLEHEAD">
            <tr>
                <td align="center">
                    <Span>Test Scripts</Span>
                </td>
                <td align="center">
                    <Span>Passed</Span>
                </td>
                <td align="center">
                    <Span>Failed</Span>
                </td>
            </tr>
            <tr align="center">
                <td>
                    <xsl:value-of select="count(TestCase)"></xsl:value-of>
                </td>
                <td>
                    <xsl:value-of select="count(TestCase[@Status='1']) + count(TestCase[@Status='4']) "></xsl:value-of>
                </td>
                <td>
                    <xsl:value-of select="count(TestCase[@Status='3']) "></xsl:value-of>
                </td>
            </tr>
        </Table>
        <BR></BR>
        <img src="_images/check-plus.jpg" onClick="expandall()" id="m1" ></img><xsl:text> </xsl:text><a href="#" onClick="expandall()">Expand All</a>
        <xsl:text> </xsl:text>
        <img src="_images/check-min.jpg" onClick="collapseall()" id="m2"></img><xsl:text> </xsl:text><a href="#" onClick="collapseall()">Collapse All</a>
        <BR></BR>
        <Div id="Batch" Style="POSITION: relative; display:block" class="TABLEHEAD">
            <Table cellSpacing="1" cellPadding="1" align="center" Class="Batch" onClick="toggleMenu('DIV{position()}.0', 'm{position()}.0')">
                <tr>
                    <td class="message">
                        <img id="m{position()}.0" src="_images/check-plus.jpg" border="0"></img>
                        <xsl:text></xsl:text><xsl:value-of select="@Desc"></xsl:value-of>
                    </td>
                    <td calss="passed">Passed -
                        <xsl:value-of select="count(TestCase/BP/Result[@Status='1'])"></xsl:value-of>
                    </td>
                    <td calss="failed">Failed -
                        <xsl:value-of select="count(TestCase/BP/Result[@Status='3']) + count(TestCase/BP/Result[@Status='0'])"></xsl:value-of>
                    </td>
                    <td calss="warnings">Warnings -
                        <xsl:value-of select="count(TestCase/BP/Result[@Status='2'])"></xsl:value-of>
                    </td>
                    <td calss="additional_information">No Run -
                        <xsl:value-of select="count(TestCase/BP/Result[@Status='4'])"></xsl:value-of>
                    </td>
                </tr>
            </Table>
        </Div>
        <Div id="DIV{position()}.0" Style="display:none">
            <xsl:apply-templates select="TestCase">
                <xsl:with-param name="BatchPosition" select="position()"/>
            </xsl:apply-templates>
        </Div>

    </xsl:template>
    <xsl:template match="TestCase">
        <xsl:param name="BatchPosition"/>
        <DIV id="TestCase" Style="LEFT: 20px; POSITION: relative; display:block">
            <TABLE cellSpacing="1" cellPadding="1" align="center" Class="TestCase" onClick="toggleMenu('DIV{concat($BatchPosition,'.',position())}', 'm{concat($BatchPosition,'.',position())}')" >
                <TR>
                    <TD class="message">
                        <img id="m{concat($BatchPosition,'.',position())}" src="_images/check-plus.jpg" border="0"></img>
                        <xsl:text> </xsl:text><xsl:value-of select="@Desc"></xsl:value-of>
                    </TD>
                    <td calss="passed">Passed -
                        <xsl:value-of select="count(BP/Result[@Status='1'])"></xsl:value-of>
                    </td>
                    <td calss="failed">Failed -
                        <xsl:value-of select="count(BP/Result[@Status='3']) + count(BP/Result[@Status='0'])"></xsl:value-of>
                    </td>
                    <td calss="warnings">Warnings -
                        <xsl:value-of select="count(BP/Result[@Status='2'])"></xsl:value-of>
                    </td>
                    <td calss="additional_information">No Run -
                        <xsl:value-of select="count(BP/Result[@Status='4'])"></xsl:value-of>
                    </td>
                </TR>
            </TABLE>
        </DIV>
        <DIV id="DIV{concat($BatchPosition,'.',position())}" Style="display:none">
            <xsl:apply-templates select="BP|ABP">
                <xsl:with-param name="TestCasePosition" select="concat($BatchPosition,'.',position())"/>
            </xsl:apply-templates>
        </DIV>

    </xsl:template>
    <xsl:template match="ABP">
        <xsl:param name="TestCasePosition"/>
        <DIV id="ABP" Style="LEFT: 40px; POSITION: relative; display:block" >
            <TABLE cellSpacing="1" cellPadding="1" align="center" Class="ABP">
                <TR>
                    <TD class="message">
                        <img id="m{concat($TestCasePosition,'.',position())}" src="_images/AggBPC.jpg" border="0"></img>
                        <xsl:text></xsl:text><xsl:value-of select="@Desc"></xsl:value-of>
                    </TD>
                </TR>
            </TABLE>
        </DIV>

    </xsl:template>
    <xsl:template match="BP">
        <xsl:param name="TestCasePosition"/>
        <DIV id="BP" Style="LEFT: 40px; POSITION: relative; display:block" >
            <TABLE cellSpacing="1" cellPadding="1" align="center" Class="BP" onClick="toggleMenu('DIV{concat($TestCasePosition,'.',position())}', 'm{concat($TestCasePosition,'.',position())}')">
                <TR>
                    <TD class="message">
                        <img id="m{concat($TestCasePosition,'.',position())}" src="_images/check-plus.jpg" border="0"></img>
                        <xsl:text></xsl:text><xsl:value-of select="@Desc"></xsl:value-of>
                    </TD>
                    <td calss="passed">Passed -
                        <xsl:value-of select="count(Result[@Status='1'])"></xsl:value-of>
                    </td>
                    <td calss="failed">Failed -
                        <xsl:value-of select="count(Result[@Status='3']) + count(Result[@Status='0'])"></xsl:value-of>
                    </td>
                    <td calss="warnings">Warnings -
                        <xsl:value-of select="count(Result[@Status='2'])"></xsl:value-of>
                    </td>
                    <td calss="additional_information">No Run -
                        <xsl:value-of select="count(Result[@Status='4'])"></xsl:value-of>
                    </td>
                </TR>
            </TABLE>
        </DIV>
        <DIV id="DIV{concat($TestCasePosition,'.',position())}" Style="display:none">
            <xsl:apply-templates select="Result">

            </xsl:apply-templates>
        </DIV>

    </xsl:template>
    <xsl:template match="Result">
        <DIV id="Result" Style="LEFT:60px; POSITION:relative; display:block" >
            <TABLE cellSpacing="1" cellPadding="1" align="Center" Class="MESSAGES">
                <xsl:element name="TR">
                    <xsl:if test="position()mod 2">
                        <xsl:attribute name="class">
                            M1
                        </xsl:attribute>
                    </xsl:if>
                    <xsl:if test="not(position()mod 2)">
                        <xsl:attribute name="class">
                            M2
                        </xsl:attribute>
                    </xsl:if>
                    <xsl:element name="TD">
                        <xsl:attribute name="class">
                            <xsl:if test="@Status='0'">
                                FAILED
                            </xsl:if>
                            <xsl:if test="@Status='3'">
                                FAILED
                            </xsl:if>
                            <xsl:if test="@Status='2'">
                                WARNINGS
                            </xsl:if>
                            <xsl:if test="@Status='1'">
                                PASSED
                            </xsl:if>
                            <xsl:if test="@Status='4'">
                                NO RUN
                            </xsl:if>
                        </xsl:attribute>
                        <xsl:text>- </xsl:text>
                        <xsl:value-of select="."></xsl:value-of>
                        <xsl:if test="@ErrorScreenShotPath">
                            <xsl:text> </xsl:text>
                            <a href="{@ErrorScreenShotPath}" target="_new">Screenshot</a>
                        </xsl:if>
                        <xsl:if test="@ScreenShotPath">
                            <xsl:text> </xsl:text>
                            <a href="{@ScreenShotPath}" target="_new">Screenshot</a>
                        </xsl:if>
                    </xsl:element>
                </xsl:element>
            </TABLE>
        </DIV>
    </xsl:template>
</xsl:stylesheet>