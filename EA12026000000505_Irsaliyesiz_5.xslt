<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2" xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2" xmlns:ccts="urn:un:unece:uncefact:documentation:2" xmlns:clm54217="urn:un:unece:uncefact:codelist:specification:54217:2001" xmlns:clm5639="urn:un:unece:uncefact:codelist:specification:5639:1988" xmlns:clm66411="urn:un:unece:uncefact:codelist:specification:66411:2001" xmlns:clmIANAMIMEMediaType="urn:un:unece:uncefact:codelist:specification:IANAMIMEMediaType:2003" xmlns:fn="http://www.w3.org/2005/xpath-functions" xmlns:link="http://www.xbrl.org/2003/linkbase" xmlns:n1="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2" xmlns:qdt="urn:oasis:names:specification:ubl:schema:xsd:QualifiedDatatypes-2" xmlns:udt="urn:un:unece:uncefact:data:specification:UnqualifiedDataTypesSchemaModule:2" xmlns:xbrldi="http://xbrl.org/2006/xbrldi" xmlns:xbrli="http://www.xbrl.org/2003/instance" xmlns:xdt="http://www.w3.org/2005/xpath-datatypes" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.0" exclude-result-prefixes="cac cbc ccts clm54217 clm5639 clm66411 clmIANAMIMEMediaType fn link n1 qdt udt xbrldi xbrli xdt xlink xs xsd xsi">
	<xsl:decimal-format name="european" decimal-separator="," grouping-separator="." NaN="" />
	<xsl:output version="4.0" method="html" indent="no" encoding="UTF-8" doctype-public="-//W3C//DTD HTML 4.01 Transitional//EN" doctype-system="http://www.w3.org/TR/html4/loose.dtd" />
	<xsl:param name="SV_OutputFormat" select="'HTML'" />
	<xsl:variable name="XML" select="/" />
	<xsl:variable name="parabirimilist" select="' TRL TRY '" />
	<xsl:template match="/">
		<html>
			<head>
			<meta http-equiv="X-UA-Compatible" content="IE=edge"/>
<script type="text/javascript">
                   <![CDATA[var QRCode;!function(){function a(a){this.mode=c.MODE_8BIT_BYTE,this.data=a,this.parsedData=[];for(var b=[],d=0,e=this.data.length;e>d;d++){var f=this.data.charCodeAt(d);f>65536?(b[0]=240|(1835008&f)>>>18,b[1]=128|(258048&f)>>>12,b[2]=128|(4032&f)>>>6,b[3]=128|63&f):f>2048?(b[0]=224|(61440&f)>>>12,b[1]=128|(4032&f)>>>6,b[2]=128|63&f):f>128?(b[0]=192|(1984&f)>>>6,b[1]=128|63&f):b[0]=f,this.parsedData=this.parsedData.concat(b)}this.parsedData.length!=this.data.length&&(this.parsedData.unshift(191),this.parsedData.unshift(187),this.parsedData.unshift(239))}function b(a,b){this.typeNumber=a,this.errorCorrectLevel=b,this.modules=null,this.moduleCount=0,this.dataCache=null,this.dataList=[]}function i(a,b){if(void 0==a.length)throw new Error(a.length+"/"+b);for(var c=0;c<a.length&&0==a[c];)c++;this.num=new Array(a.length-c+b);for(var d=0;d<a.length-c;d++)this.num[d]=a[d+c]}function j(a,b){this.totalCount=a,this.dataCount=b}function k(){this.buffer=[],this.length=0}function m(){return"undefined"!=typeof CanvasRenderingContext2D}function n(){var a=!1,b=navigator.userAgent;return/android/i.test(b)&&(a=!0,aMat=b.toString().match(/android ([0-9]\.[0-9])/i),aMat&&aMat[1]&&(a=parseFloat(aMat[1]))),a}function r(a,b){for(var c=1,e=s(a),f=0,g=l.length;g>=f;f++){var h=0;switch(b){case d.L:h=l[f][0];break;case d.M:h=l[f][1];break;case d.Q:h=l[f][2];break;case d.H:h=l[f][3]}if(h>=e)break;c++}if(c>l.length)throw new Error("Too long data");return c}function s(a){var b=encodeURI(a).toString().replace(/\%[0-9a-fA-F]{2}/g,"a");return b.length+(b.length!=a?3:0)}a.prototype={getLength:function(){return this.parsedData.length},write:function(a){for(var b=0,c=this.parsedData.length;c>b;b++)a.put(this.parsedData[b],8)}},b.prototype={addData:function(b){var c=new a(b);this.dataList.push(c),this.dataCache=null},isDark:function(a,b){if(0>a||this.moduleCount<=a||0>b||this.moduleCount<=b)throw new Error(a+","+b);return this.modules[a][b]},getModuleCount:function(){return this.moduleCount},make:function(){this.makeImpl(!1,this.getBestMaskPattern())},makeImpl:function(a,c){this.moduleCount=4*this.typeNumber+17,this.modules=new Array(this.moduleCount);for(var d=0;d<this.moduleCount;d++){this.modules[d]=new Array(this.moduleCount);for(var e=0;e<this.moduleCount;e++)this.modules[d][e]=null}this.setupPositionProbePattern(0,0),this.setupPositionProbePattern(this.moduleCount-7,0),this.setupPositionProbePattern(0,this.moduleCount-7),this.setupPositionAdjustPattern(),this.setupTimingPattern(),this.setupTypeInfo(a,c),this.typeNumber>=7&&this.setupTypeNumber(a),null==this.dataCache&&(this.dataCache=b.createData(this.typeNumber,this.errorCorrectLevel,this.dataList)),this.mapData(this.dataCache,c)},setupPositionProbePattern:function(a,b){for(var c=-1;7>=c;c++)if(!(-1>=a+c||this.moduleCount<=a+c))for(var d=-1;7>=d;d++)-1>=b+d||this.moduleCount<=b+d||(this.modules[a+c][b+d]=c>=0&&6>=c&&(0==d||6==d)||d>=0&&6>=d&&(0==c||6==c)||c>=2&&4>=c&&d>=2&&4>=d?!0:!1)},getBestMaskPattern:function(){for(var a=0,b=0,c=0;8>c;c++){this.makeImpl(!0,c);var d=f.getLostPoint(this);(0==c||a>d)&&(a=d,b=c)}return b},createMovieClip:function(a,b,c){var d=a.createEmptyMovieClip(b,c),e=1;this.make();for(var f=0;f<this.modules.length;f++)for(var g=f*e,h=0;h<this.modules[f].length;h++){var i=h*e,j=this.modules[f][h];j&&(d.beginFill(0,100),d.moveTo(i,g),d.lineTo(i+e,g),d.lineTo(i+e,g+e),d.lineTo(i,g+e),d.endFill())}return d},setupTimingPattern:function(){for(var a=8;a<this.moduleCount-8;a++)null==this.modules[a][6]&&(this.modules[a][6]=0==a%2);for(var b=8;b<this.moduleCount-8;b++)null==this.modules[6][b]&&(this.modules[6][b]=0==b%2)},setupPositionAdjustPattern:function(){for(var a=f.getPatternPosition(this.typeNumber),b=0;b<a.length;b++)for(var c=0;c<a.length;c++){var d=a[b],e=a[c];if(null==this.modules[d][e])for(var g=-2;2>=g;g++)for(var h=-2;2>=h;h++)this.modules[d+g][e+h]=-2==g||2==g||-2==h||2==h||0==g&&0==h?!0:!1}},setupTypeNumber:function(a){for(var b=f.getBCHTypeNumber(this.typeNumber),c=0;18>c;c++){var d=!a&&1==(1&b>>c);this.modules[Math.floor(c/3)][c%3+this.moduleCount-8-3]=d}for(var c=0;18>c;c++){var d=!a&&1==(1&b>>c);this.modules[c%3+this.moduleCount-8-3][Math.floor(c/3)]=d}},setupTypeInfo:function(a,b){for(var c=this.errorCorrectLevel<<3|b,d=f.getBCHTypeInfo(c),e=0;15>e;e++){var g=!a&&1==(1&d>>e);6>e?this.modules[e][8]=g:8>e?this.modules[e+1][8]=g:this.modules[this.moduleCount-15+e][8]=g}for(var e=0;15>e;e++){var g=!a&&1==(1&d>>e);8>e?this.modules[8][this.moduleCount-e-1]=g:9>e?this.modules[8][15-e-1+1]=g:this.modules[8][15-e-1]=g}this.modules[this.moduleCount-8][8]=!a},mapData:function(a,b){for(var c=-1,d=this.moduleCount-1,e=7,g=0,h=this.moduleCount-1;h>0;h-=2)for(6==h&&h--;;){for(var i=0;2>i;i++)if(null==this.modules[d][h-i]){var j=!1;g<a.length&&(j=1==(1&a[g]>>>e));var k=f.getMask(b,d,h-i);k&&(j=!j),this.modules[d][h-i]=j,e--,-1==e&&(g++,e=7)}if(d+=c,0>d||this.moduleCount<=d){d-=c,c=-c;break}}}},b.PAD0=236,b.PAD1=17,b.createData=function(a,c,d){for(var e=j.getRSBlocks(a,c),g=new k,h=0;h<d.length;h++){var i=d[h];g.put(i.mode,4),g.put(i.getLength(),f.getLengthInBits(i.mode,a)),i.write(g)}for(var l=0,h=0;h<e.length;h++)l+=e[h].dataCount;if(g.getLengthInBits()>8*l)throw new Error("code length overflow. ("+g.getLengthInBits()+">"+8*l+")");for(g.getLengthInBits()+4<=8*l&&g.put(0,4);0!=g.getLengthInBits()%8;)g.putBit(!1);for(;;){if(g.getLengthInBits()>=8*l)break;if(g.put(b.PAD0,8),g.getLengthInBits()>=8*l)break;g.put(b.PAD1,8)}return b.createBytes(g,e)},b.createBytes=function(a,b){for(var c=0,d=0,e=0,g=new Array(b.length),h=new Array(b.length),j=0;j<b.length;j++){var k=b[j].dataCount,l=b[j].totalCount-k;d=Math.max(d,k),e=Math.max(e,l),g[j]=new Array(k);for(var m=0;m<g[j].length;m++)g[j][m]=255&a.buffer[m+c];c+=k;var n=f.getErrorCorrectPolynomial(l),o=new i(g[j],n.getLength()-1),p=o.mod(n);h[j]=new Array(n.getLength()-1);for(var m=0;m<h[j].length;m++){var q=m+p.getLength()-h[j].length;h[j][m]=q>=0?p.get(q):0}}for(var r=0,m=0;m<b.length;m++)r+=b[m].totalCount;for(var s=new Array(r),t=0,m=0;d>m;m++)for(var j=0;j<b.length;j++)m<g[j].length&&(s[t++]=g[j][m]);for(var m=0;e>m;m++)for(var j=0;j<b.length;j++)m<h[j].length&&(s[t++]=h[j][m]);return s};for(var c={MODE_NUMBER:1,MODE_ALPHA_NUM:2,MODE_8BIT_BYTE:4,MODE_KANJI:8},d={L:1,M:0,Q:3,H:2},e={PATTERN000:0,PATTERN001:1,PATTERN010:2,PATTERN011:3,PATTERN100:4,PATTERN101:5,PATTERN110:6,PATTERN111:7},f={PATTERN_POSITION_TABLE:[[],[6,18],[6,22],[6,26],[6,30],[6,34],[6,22,38],[6,24,42],[6,26,46],[6,28,50],[6,30,54],[6,32,58],[6,34,62],[6,26,46,66],[6,26,48,70],[6,26,50,74],[6,30,54,78],[6,30,56,82],[6,30,58,86],[6,34,62,90],[6,28,50,72,94],[6,26,50,74,98],[6,30,54,78,102],[6,28,54,80,106],[6,32,58,84,110],[6,30,58,86,114],[6,34,62,90,118],[6,26,50,74,98,122],[6,30,54,78,102,126],[6,26,52,78,104,130],[6,30,56,82,108,134],[6,34,60,86,112,138],[6,30,58,86,114,142],[6,34,62,90,118,146],[6,30,54,78,102,126,150],[6,24,50,76,102,128,154],[6,28,54,80,106,132,158],[6,32,58,84,110,136,162],[6,26,54,82,110,138,166],[6,30,58,86,114,142,170]],G15:1335,G18:7973,G15_MASK:21522,getBCHTypeInfo:function(a){for(var b=a<<10;f.getBCHDigit(b)-f.getBCHDigit(f.G15)>=0;)b^=f.G15<<f.getBCHDigit(b)-f.getBCHDigit(f.G15);return(a<<10|b)^f.G15_MASK},getBCHTypeNumber:function(a){for(var b=a<<12;f.getBCHDigit(b)-f.getBCHDigit(f.G18)>=0;)b^=f.G18<<f.getBCHDigit(b)-f.getBCHDigit(f.G18);return a<<12|b},getBCHDigit:function(a){for(var b=0;0!=a;)b++,a>>>=1;return b},getPatternPosition:function(a){return f.PATTERN_POSITION_TABLE[a-1]},getMask:function(a,b,c){switch(a){case e.PATTERN000:return 0==(b+c)%2;case e.PATTERN001:return 0==b%2;case e.PATTERN010:return 0==c%3;case e.PATTERN011:return 0==(b+c)%3;case e.PATTERN100:return 0==(Math.floor(b/2)+Math.floor(c/3))%2;case e.PATTERN101:return 0==b*c%2+b*c%3;case e.PATTERN110:return 0==(b*c%2+b*c%3)%2;case e.PATTERN111:return 0==(b*c%3+(b+c)%2)%2;default:throw new Error("bad maskPattern:"+a)}},getErrorCorrectPolynomial:function(a){for(var b=new i([1],0),c=0;a>c;c++)b=b.multiply(new i([1,g.gexp(c)],0));return b},getLengthInBits:function(a,b){if(b>=1&&10>b)switch(a){case c.MODE_NUMBER:return 10;case c.MODE_ALPHA_NUM:return 9;case c.MODE_8BIT_BYTE:return 8;case c.MODE_KANJI:return 8;default:throw new Error("mode:"+a)}else if(27>b)switch(a){case c.MODE_NUMBER:return 12;case c.MODE_ALPHA_NUM:return 11;case c.MODE_8BIT_BYTE:return 16;case c.MODE_KANJI:return 10;default:throw new Error("mode:"+a)}else{if(!(41>b))throw new Error("type:"+b);switch(a){case c.MODE_NUMBER:return 14;case c.MODE_ALPHA_NUM:return 13;case c.MODE_8BIT_BYTE:return 16;case c.MODE_KANJI:return 12;default:throw new Error("mode:"+a)}}},getLostPoint:function(a){for(var b=a.getModuleCount(),c=0,d=0;b>d;d++)for(var e=0;b>e;e++){for(var f=0,g=a.isDark(d,e),h=-1;1>=h;h++)if(!(0>d+h||d+h>=b))for(var i=-1;1>=i;i++)0>e+i||e+i>=b||(0!=h||0!=i)&&g==a.isDark(d+h,e+i)&&f++;f>5&&(c+=3+f-5)}for(var d=0;b-1>d;d++)for(var e=0;b-1>e;e++){var j=0;a.isDark(d,e)&&j++,a.isDark(d+1,e)&&j++,a.isDark(d,e+1)&&j++,a.isDark(d+1,e+1)&&j++,(0==j||4==j)&&(c+=3)}for(var d=0;b>d;d++)for(var e=0;b-6>e;e++)a.isDark(d,e)&&!a.isDark(d,e+1)&&a.isDark(d,e+2)&&a.isDark(d,e+3)&&a.isDark(d,e+4)&&!a.isDark(d,e+5)&&a.isDark(d,e+6)&&(c+=40);for(var e=0;b>e;e++)for(var d=0;b-6>d;d++)a.isDark(d,e)&&!a.isDark(d+1,e)&&a.isDark(d+2,e)&&a.isDark(d+3,e)&&a.isDark(d+4,e)&&!a.isDark(d+5,e)&&a.isDark(d+6,e)&&(c+=40);for(var k=0,e=0;b>e;e++)for(var d=0;b>d;d++)a.isDark(d,e)&&k++;var l=Math.abs(100*k/b/b-50)/5;return c+=10*l}},g={glog:function(a){if(1>a)throw new Error("glog("+a+")");return g.LOG_TABLE[a]},gexp:function(a){for(;0>a;)a+=255;for(;a>=256;)a-=255;return g.EXP_TABLE[a]},EXP_TABLE:new Array(256),LOG_TABLE:new Array(256)},h=0;8>h;h++)g.EXP_TABLE[h]=1<<h;for(var h=8;256>h;h++)g.EXP_TABLE[h]=g.EXP_TABLE[h-4]^g.EXP_TABLE[h-5]^g.EXP_TABLE[h-6]^g.EXP_TABLE[h-8];for(var h=0;255>h;h++)g.LOG_TABLE[g.EXP_TABLE[h]]=h;i.prototype={get:function(a){return this.num[a]},getLength:function(){return this.num.length},multiply:function(a){for(var b=new Array(this.getLength()+a.getLength()-1),c=0;c<this.getLength();c++)for(var d=0;d<a.getLength();d++)b[c+d]^=g.gexp(g.glog(this.get(c))+g.glog(a.get(d)));return new i(b,0)},mod:function(a){if(this.getLength()-a.getLength()<0)return this;for(var b=g.glog(this.get(0))-g.glog(a.get(0)),c=new Array(this.getLength()),d=0;d<this.getLength();d++)c[d]=this.get(d);for(var d=0;d<a.getLength();d++)c[d]^=g.gexp(g.glog(a.get(d))+b);return new i(c,0).mod(a)}},j.RS_BLOCK_TABLE=[[1,26,19],[1,26,16],[1,26,13],[1,26,9],[1,44,34],[1,44,28],[1,44,22],[1,44,16],[1,70,55],[1,70,44],[2,35,17],[2,35,13],[1,100,80],[2,50,32],[2,50,24],[4,25,9],[1,134,108],[2,67,43],[2,33,15,2,34,16],[2,33,11,2,34,12],[2,86,68],[4,43,27],[4,43,19],[4,43,15],[2,98,78],[4,49,31],[2,32,14,4,33,15],[4,39,13,1,40,14],[2,121,97],[2,60,38,2,61,39],[4,40,18,2,41,19],[4,40,14,2,41,15],[2,146,116],[3,58,36,2,59,37],[4,36,16,4,37,17],[4,36,12,4,37,13],[2,86,68,2,87,69],[4,69,43,1,70,44],[6,43,19,2,44,20],[6,43,15,2,44,16],[4,101,81],[1,80,50,4,81,51],[4,50,22,4,51,23],[3,36,12,8,37,13],[2,116,92,2,117,93],[6,58,36,2,59,37],[4,46,20,6,47,21],[7,42,14,4,43,15],[4,133,107],[8,59,37,1,60,38],[8,44,20,4,45,21],[12,33,11,4,34,12],[3,145,115,1,146,116],[4,64,40,5,65,41],[11,36,16,5,37,17],[11,36,12,5,37,13],[5,109,87,1,110,88],[5,65,41,5,66,42],[5,54,24,7,55,25],[11,36,12],[5,122,98,1,123,99],[7,73,45,3,74,46],[15,43,19,2,44,20],[3,45,15,13,46,16],[1,135,107,5,136,108],[10,74,46,1,75,47],[1,50,22,15,51,23],[2,42,14,17,43,15],[5,150,120,1,151,121],[9,69,43,4,70,44],[17,50,22,1,51,23],[2,42,14,19,43,15],[3,141,113,4,142,114],[3,70,44,11,71,45],[17,47,21,4,48,22],[9,39,13,16,40,14],[3,135,107,5,136,108],[3,67,41,13,68,42],[15,54,24,5,55,25],[15,43,15,10,44,16],[4,144,116,4,145,117],[17,68,42],[17,50,22,6,51,23],[19,46,16,6,47,17],[2,139,111,7,140,112],[17,74,46],[7,54,24,16,55,25],[34,37,13],[4,151,121,5,152,122],[4,75,47,14,76,48],[11,54,24,14,55,25],[16,45,15,14,46,16],[6,147,117,4,148,118],[6,73,45,14,74,46],[11,54,24,16,55,25],[30,46,16,2,47,17],[8,132,106,4,133,107],[8,75,47,13,76,48],[7,54,24,22,55,25],[22,45,15,13,46,16],[10,142,114,2,143,115],[19,74,46,4,75,47],[28,50,22,6,51,23],[33,46,16,4,47,17],[8,152,122,4,153,123],[22,73,45,3,74,46],[8,53,23,26,54,24],[12,45,15,28,46,16],[3,147,117,10,148,118],[3,73,45,23,74,46],[4,54,24,31,55,25],[11,45,15,31,46,16],[7,146,116,7,147,117],[21,73,45,7,74,46],[1,53,23,37,54,24],[19,45,15,26,46,16],[5,145,115,10,146,116],[19,75,47,10,76,48],[15,54,24,25,55,25],[23,45,15,25,46,16],[13,145,115,3,146,116],[2,74,46,29,75,47],[42,54,24,1,55,25],[23,45,15,28,46,16],[17,145,115],[10,74,46,23,75,47],[10,54,24,35,55,25],[19,45,15,35,46,16],[17,145,115,1,146,116],[14,74,46,21,75,47],[29,54,24,19,55,25],[11,45,15,46,46,16],[13,145,115,6,146,116],[14,74,46,23,75,47],[44,54,24,7,55,25],[59,46,16,1,47,17],[12,151,121,7,152,122],[12,75,47,26,76,48],[39,54,24,14,55,25],[22,45,15,41,46,16],[6,151,121,14,152,122],[6,75,47,34,76,48],[46,54,24,10,55,25],[2,45,15,64,46,16],[17,152,122,4,153,123],[29,74,46,14,75,47],[49,54,24,10,55,25],[24,45,15,46,46,16],[4,152,122,18,153,123],[13,74,46,32,75,47],[48,54,24,14,55,25],[42,45,15,32,46,16],[20,147,117,4,148,118],[40,75,47,7,76,48],[43,54,24,22,55,25],[10,45,15,67,46,16],[19,148,118,6,149,119],[18,75,47,31,76,48],[34,54,24,34,55,25],[20,45,15,61,46,16]],j.getRSBlocks=function(a,b){var c=j.getRsBlockTable(a,b);if(void 0==c)throw new Error("bad rs block @ typeNumber:"+a+"/errorCorrectLevel:"+b);for(var d=c.length/3,e=[],f=0;d>f;f++)for(var g=c[3*f+0],h=c[3*f+1],i=c[3*f+2],k=0;g>k;k++)e.push(new j(h,i));return e},j.getRsBlockTable=function(a,b){switch(b){case d.L:return j.RS_BLOCK_TABLE[4*(a-1)+0];case d.M:return j.RS_BLOCK_TABLE[4*(a-1)+1];case d.Q:return j.RS_BLOCK_TABLE[4*(a-1)+2];case d.H:return j.RS_BLOCK_TABLE[4*(a-1)+3];default:return void 0}},k.prototype={get:function(a){var b=Math.floor(a/8);return 1==(1&this.buffer[b]>>>7-a%8)},put:function(a,b){for(var c=0;b>c;c++)this.putBit(1==(1&a>>>b-c-1))},getLengthInBits:function(){return this.length},putBit:function(a){var b=Math.floor(this.length/8);this.buffer.length<=b&&this.buffer.push(0),a&&(this.buffer[b]|=128>>>this.length%8),this.length++}};var l=[[17,14,11,7],[32,26,20,14],[53,42,32,24],[78,62,46,34],[106,84,60,44],[134,106,74,58],[154,122,86,64],[192,152,108,84],[230,180,130,98],[271,213,151,119],[321,251,177,137],[367,287,203,155],[425,331,241,177],[458,362,258,194],[520,412,292,220],[586,450,322,250],[644,504,364,280],[718,560,394,310],[792,624,442,338],[858,666,482,382],[929,711,509,403],[1003,779,565,439],[1091,857,611,461],[1171,911,661,511],[1273,997,715,535],[1367,1059,751,593],[1465,1125,805,625],[1528,1190,868,658],[1628,1264,908,698],[1732,1370,982,742],[1840,1452,1030,790],[1952,1538,1112,842],[2068,1628,1168,898],[2188,1722,1228,958],[2303,1809,1283,983],[2431,1911,1351,1051],[2563,1989,1423,1093],[2699,2099,1499,1139],[2809,2213,1579,1219],[2953,2331,1663,1273]],o=function(){var a=function(a,b){this._el=a,this._htOption=b};return a.prototype.draw=function(a){function g(a,b){var c=document.createElementNS("http://www.w3.org/2000/svg",a);for(var d in b)b.hasOwnProperty(d)&&c.setAttribute(d,b[d]);return c}var b=this._htOption,c=this._el,d=a.getModuleCount();Math.floor(b.width/d),Math.floor(b.height/d),this.clear();var h=g("svg",{viewBox:"0 0 "+String(d)+" "+String(d),width:"100%",height:"100%",fill:b.colorLight});h.setAttributeNS("http://www.w3.org/2000/xmlns/","xmlns:xlink","http://www.w3.org/1999/xlink"),c.appendChild(h),h.appendChild(g("rect",{fill:b.colorDark,width:"1",height:"1",id:"template"}));for(var i=0;d>i;i++)for(var j=0;d>j;j++)if(a.isDark(i,j)){var k=g("use",{x:String(i),y:String(j)});k.setAttributeNS("http://www.w3.org/1999/xlink","href","#template"),h.appendChild(k)}},a.prototype.clear=function(){for(;this._el.hasChildNodes();)this._el.removeChild(this._el.lastChild)},a}(),p="svg"===document.documentElement.tagName.toLowerCase(),q=p?o:m()?function(){function a(){this._elImage.src=this._elCanvas.toDataURL("image/png"),this._elImage.style.display="block",this._elCanvas.style.display="none"}function d(a,b){var c=this;if(c._fFail=b,c._fSuccess=a,null===c._bSupportDataURI){var d=document.createElement("img"),e=function(){c._bSupportDataURI=!1,c._fFail&&_fFail.call(c)},f=function(){c._bSupportDataURI=!0,c._fSuccess&&c._fSuccess.call(c)};return d.onabort=e,d.onerror=e,d.onload=f,d.src="data:image/gif;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==",void 0}c._bSupportDataURI===!0&&c._fSuccess?c._fSuccess.call(c):c._bSupportDataURI===!1&&c._fFail&&c._fFail.call(c)}if(this._android&&this._android<=2.1){var b=1/window.devicePixelRatio,c=CanvasRenderingContext2D.prototype.drawImage;CanvasRenderingContext2D.prototype.drawImage=function(a,d,e,f,g,h,i,j){if("nodeName"in a&&/img/i.test(a.nodeName))for(var l=arguments.length-1;l>=1;l--)arguments[l]=arguments[l]*b;else"undefined"==typeof j&&(arguments[1]*=b,arguments[2]*=b,arguments[3]*=b,arguments[4]*=b);c.apply(this,arguments)}}var e=function(a,b){this._bIsPainted=!1,this._android=n(),this._htOption=b,this._elCanvas=document.createElement("canvas"),this._elCanvas.width=b.width,this._elCanvas.height=b.height,a.appendChild(this._elCanvas),this._el=a,this._oContext=this._elCanvas.getContext("2d"),this._bIsPainted=!1,this._elImage=document.createElement("img"),this._elImage.style.display="none",this._el.appendChild(this._elImage),this._bSupportDataURI=null};return e.prototype.draw=function(a){var b=this._elImage,c=this._oContext,d=this._htOption,e=a.getModuleCount(),f=d.width/e,g=d.height/e,h=Math.round(f),i=Math.round(g);b.style.display="none",this.clear();for(var j=0;e>j;j++)for(var k=0;e>k;k++){var l=a.isDark(j,k),m=k*f,n=j*g;c.strokeStyle=l?d.colorDark:d.colorLight,c.lineWidth=1,c.fillStyle=l?d.colorDark:d.colorLight,c.fillRect(m,n,f,g),c.strokeRect(Math.floor(m)+.5,Math.floor(n)+.5,h,i),c.strokeRect(Math.ceil(m)-.5,Math.ceil(n)-.5,h,i)}this._bIsPainted=!0},e.prototype.makeImage=function(){this._bIsPainted&&d.call(this,a)},e.prototype.isPainted=function(){return this._bIsPainted},e.prototype.clear=function(){this._oContext.clearRect(0,0,this._elCanvas.width,this._elCanvas.height),this._bIsPainted=!1},e.prototype.round=function(a){return a?Math.floor(1e3*a)/1e3:a},e}():function(){var a=function(a,b){this._el=a,this._htOption=b};return a.prototype.draw=function(a){for(var b=this._htOption,c=this._el,d=a.getModuleCount(),e=Math.floor(b.width/d),f=Math.floor(b.height/d),g=['<table style="border:0;border-collapse:collapse;">'],h=0;d>h;h++){g.push("<tr>");for(var i=0;d>i;i++)g.push('<td style="border:0;border-collapse:collapse;padding:0;margin:0;width:'+e+"px;height:"+f+"px;background-color:"+(a.isDark(h,i)?b.colorDark:b.colorLight)+';"></td>');g.push("</tr>")}g.push("</table>"),c.innerHTML=g.join("");var j=c.childNodes[0],k=(b.width-j.offsetWidth)/2,l=(b.height-j.offsetHeight)/2;k>0&&l>0&&(j.style.margin=l+"px "+k+"px")},a.prototype.clear=function(){this._el.innerHTML=""},a}();QRCode=function(a,b){if(this._htOption={width:256,height:256,typeNumber:4,colorDark:"#000000",colorLight:"#ffffff",correctLevel:d.H},"string"==typeof b&&(b={text:b}),b)for(var c in b)this._htOption[c]=b[c];"string"==typeof a&&(a=document.getElementById(a)),this._android=n(),this._el=a,this._oQRCode=null,this._oDrawing=new q(this._el,this._htOption),this._htOption.text&&this.makeCode(this._htOption.text)},QRCode.prototype.makeCode=function(a){this._oQRCode=new b(r(a,this._htOption.correctLevel),this._htOption.correctLevel),this._oQRCode.addData(a),this._oQRCode.make(),this._el.title=a,this._oDrawing.draw(this._oQRCode),this.makeImage()},QRCode.prototype.makeImage=function(){"function"==typeof this._oDrawing.makeImage&&(!this._android||this._android>=3)&&this._oDrawing.makeImage()},QRCode.prototype.clear=function(){this._oDrawing.clear()},QRCode.CorrectLevel=d}();]]>	           	
				</script>
				<title />
				<style type="text/css">  body { background-color: #FFFFFF; font-family: 'Tahoma', "Times New Roman", Times, serif; font-size: 11px; color: #666666; } h1, h2 { padding-bottom: 3px; padding-top: 3px; margin-bottom: 5px; text-transform: uppercase; font-family: Arial, Helvetica, sans-serif; } h1 { font-size: 1.4em; text-transform:none; } h2 { font-size: 1em; color: brown; } h3 { font-size: 1em; color: #333333; text-align: justify; margin: 0; padding: 0; } h4 { font-size: 1.1em; font-style: bold; font-family: Arial, Helvetica, sans-serif; color: #000000; margin: 0; padding: 0; } hr { height:2px; color: #000000; background-color: #000000; border-bottom: 1px solid #000000; } p, ul, ol { margin-top: 1.5em; } ul, ol { margin-left: 3em; } blockquote { margin-left: 3em; margin-right: 3em; font-style: italic; } a { text-decoration: none; color: #70A300; } a:hover { border: none; color: #70A300; } #despatchTable > tbody:first-of-type > tr:nth-child(n+6) { display:none; } #despatchTable { border-collapse:collapse; font-size:11px; float:right; border-color:gray; } #ettnTable { border-collapse:collapse; font-size:11px; border-color:gray; } #customerPartyTable { border-width: 0px; border-spacing:; border-style: inset; border-color: gray; border-collapse: collapse; background-color: } #customerIDTable { border-width: 2px; border-spacing:; border-style: inset; border-color: gray; border-collapse: collapse; background-color: } #customerIDTableTd { border-width: 2px; border-spacing:; border-style: inset; border-color: gray; border-collapse: collapse; background-color: } #lineTable { border-width:2px; border-spacing:; border-style: inset; border-color: black; border-collapse: collapse; background-color:; } #lineTableTd { border-width: 1px; padding: 1px; border-style: inset; border-color: black; background-color: white; } #lineTableTr { border-width: 1px; padding: 0px; border-style: inset; border-color: black; background-color: white; -moz-border-radius:; } #lineTableDummyTd { border-width: 1px; border-color:white; padding: 1px; border-style: inset; border-color: black; background-color: white; } #lineTableBudgetTd { border-width: 2px; border-spacing:0px; padding: 1px; border-style: inset; border-color: black; background-color: white; -moz-border-radius:; } #notesTable { border-width: 2px; border-spacing:; border-style: inset; border-color: black; border-collapse: collapse; background-color: } #notesTableTd { border-width: 0px; border-spacing:; border-style: inset; border-color: black; border-collapse: collapse; background-color: } table { border-spacing:0px; } #budgetContainerTable { border-width: 0px; border-spacing: 0px; border-style: inset; border-color: black; border-collapse: collapse; background-color:; } td { border-color:gray; }</style>
				<title>e-Arşiv</title>
			</head>
			<body style="margin-left=0.6in; margin-right=0.6in; margin-top=0.79in; margin-bottom=0.79in">
				<xsl:for-each select="$XML">
					<table style="border-color:blue; " border="0" cellspacing="0px" width="800" cellpadding="0px">
						<tbody>
							<tr valign="top">
								<td width="60%" colspan="2">
									<br />
									<table align="center" border="0" width="100%" style="color:#03579C">
										<tbody>
											<tr align="center">
												<td colspan="4" style="font-weight:bold;font-size:14px;">
													<xsl:text>Ant Teknik Cihazlar Pazarlama ve Dış Ticaret Anonim Şirketi</xsl:text>
												</td>
											</tr>
											<tr align="center">
												<td colspan="4" style="font-weight:bold;font-size:12px;">
													<xsl:text> </xsl:text>
												</td>
											</tr>
											<tr align="left">
												<td align="left">
													<xsl:text>MERKEZ:Burhaniye Mah.</xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>ŞUBE:Kızılcaşar Mah.</xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>ŞUBE:Tepekule Mah.2084/7 Sk.</xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>ŞUBE:Yeni Mah. 87071 Sok. </xsl:text>
												</td>
											</tr>
											<tr align="left">
												<td align="left">
													<xsl:text>Beybostanı Sk. No:45 </xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>23 Nisan Caddesi No:146 </xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>İrem Sit. Ömerağa Apt. No:4 </xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text> Bozkurtlar Rezidans No:52/19</xsl:text>
												</td>
											</tr>
											<tr align="left">
												<td align="left">
													<xsl:text>34676 Beylerbeyi</xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>06830 Gölbaşı</xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>Kat:3 D:12 Bayraklı</xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>01200 Seyhan</xsl:text>
												</td>
											</tr>
											<tr align="left">
												<td align="left">
													<xsl:text>İstanbul/TÜRKİYE</xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>Ankara /TÜRKİYE</xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>İzmir/TÜRKİYE</xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>Adana/TÜRKİYE</xsl:text>
												</td>
											</tr>
											<tr align="left">
												<td align="left">
													<xsl:text>Tel   :(0216) 422 67 00 pbx</xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>Tel   :(0312) 472 87 40 pbx</xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>Tel   :(0232) 371 62 00 pbx</xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>Tel: (0322) 457 4501 pbx</xsl:text>
												</td>
											</tr>
											<tr align="left">
												<td align="left">
													<xsl:text>Faks:(0216) 422 39 54</xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>Faks:(0312) 472 87 42</xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>Faks:(0232) 371 62 10</xsl:text>
												</td>
												<td align="left" style="border-left:1px solid">
													<xsl:text>Faks:(0322) 457 45 02</xsl:text>
												</td>
											</tr>
											<tr align="center">
												<td colspan="4" style="font-weight:bold;font-size:12px;">
													<xsl:text> </xsl:text>
												</td>
											</tr>
											<tr>
												<td colspan="4" align="center">
													<xsl:text>E-mail : antteknik@antteknik.com  Web: http://www.antteknik.com</xsl:text>
												</td>
											</tr>
											<tr>
												<td colspan="4" align="center">
													<xsl:text>Tic. Sic. : 413499-361081   YENİKAPI V.D. : 0701234931</xsl:text>
												</td>
											</tr>
											<!-- <tr align="left"> <xsl:for-each select="n1:Invoice"> <xsl:for-each select="cac:AccountingSupplierParty"> <xsl:for-each select="cac:Party"> <td align="left"> <xsl:if test="cac:PartyName"> <xsl:value-of select="cac:PartyName/cbc:Name"/> <br/> </xsl:if> <xsl:for-each select="cac:Person"> <xsl:for-each select="cbc:Title"> <xsl:apply-templates/> <span> <<xsl:text>&#160;</xsl:text> </span> </xsl:for-each> <xsl:for-each select="cbc:FirstName"> <xsl:apply-templates/> <span> <xsl:text>&#160;</xsl:text> </span> </xsl:for-each> <xsl:for-each select="cbc:MiddleName"> <xsl:apply-templates/> <span> <xsl:text>&#160;</xsl:text> </span> </xsl:for-each> <xsl:for-each select="cbc:FamilyName"> <xsl:apply-templates/> <span> <xsl:text>&#160;</xsl:text> </span> </xsl:for-each> <xsl:for-each select="cbc:NameSuffix"> <xsl:apply-templates/> </xsl:for-each> </xsl:for-each> </td> </xsl:for-each> </xsl:for-each> </xsl:for-each> </tr> <tr align="left"> <xsl:for-each select="n1:Invoice"> <xsl:for-each select="cac:AccountingSupplierParty"> <xsl:for-each select="cac:Party"> <td align="left"> <xsl:for-each select="cac:PostalAddress"> <xsl:for-each select="cbc:StreetName"> <xsl:apply-templates/> <span> <xsl:text>&#160;</xsl:text> </span> </xsl:for-each> <xsl:for-each select="cbc:BuildingName"> <xsl:apply-templates/> </xsl:for-each> <xsl:if test="cbc:BuildingNumber"> <span> <xsl:text> No:</xsl:text> </span> <xsl:for-each select="cbc:BuildingNumber"> <xsl:apply-templates/> </xsl:for-each> <span> <xsl:text>&#160;</xsl:text> </span> </xsl:if> <br/> <xsl:for-each select="cbc:PostalZone"> <xsl:apply-templates/> <span> <xsl:text>&#160;</xsl:text> </span> </xsl:for-each> <xsl:for-each select="cbc:CitySubdivisionName"> <xsl:apply-templates/> </xsl:for-each> <span> <xsl:text>/ </xsl:text> </span> <xsl:for-each select="cbc:CityName"> <xsl:apply-templates/> <span> <xsl:text>&#160;</xsl:text> </span> </xsl:for-each> </xsl:for-each> </td> </xsl:for-each> </xsl:for-each> </xsl:for-each> </tr> <xsl:if test="//n1:Invoice/cac:AccountingSupplierParty/cac:Party/cac:Contact/cbc:Telephone or //n1:Invoice/cac:AccountingSupplierParty/cac:Party/cac:Contact/cbc:Telefax"> <tr align="left"> <xsl:for-each select="n1:Invoice"> <xsl:for-each select="cac:AccountingSupplierParty"> <xsl:for-each select="cac:Party"> <td align="left"> <xsl:for-each select="cac:Contact"> <xsl:if test="cbc:Telephone"> <span> <xsl:text>Tel: </xsl:text> </span> <xsl:for-each select="cbc:Telephone"> <xsl:apply-templates/> </xsl:for-each> </xsl:if> <xsl:if test="cbc:Telefax"> <span> <xsl:text> Fax: </xsl:text> </span> <xsl:for-each select="cbc:Telefax"> <xsl:apply-templates/> </xsl:for-each> </xsl:if> <span> <xsl:text>&#160;</xsl:text> </span> </xsl:for-each> </td> </xsl:for-each> </xsl:for-each> </xsl:for-each> </tr> </xsl:if> <xsl:for-each select="//n1:Invoice/cac:AccountingSupplierParty/cac:Party/cbc:WebsiteURI"> <tr align="left"> <td> <xsl:text>Web Sitesi: </xsl:text> <xsl:value-of select="."/> </td> </tr> </xsl:for-each> <xsl:for-each select="//n1:Invoice/cac:AccountingSupplierParty/cac:Party/cac:Contact/cbc:ElectronicMail"> <tr align="left"> <td> <xsl:text>E-Posta: </xsl:text> <xsl:value-of select="."/> </td> </tr> </xsl:for-each> <tr align="left"> <xsl:for-each select="n1:Invoice"> <xsl:for-each select="cac:AccountingSupplierParty"> <xsl:for-each select="cac:Party"> <td align="left"> <span> <xsl:text>Vergi Dairesi: </xsl:text> </span> <xsl:for-each select="cac:PartyTaxScheme"> <xsl:for-each select="cac:TaxScheme"> <xsl:for-each select="cbc:Name"> <xsl:apply-templates/> </xsl:for-each> </xsl:for-each> <span> <xsl:text>&#160; </xsl:text> </span> </xsl:for-each> </td> </xsl:for-each> </xsl:for-each> </xsl:for-each> </tr> <xsl:for-each select="//n1:Invoice/cac:AccountingSupplierParty/cac:Party/cac:PartyIdentification"> <tr align="left"> <td> <xsl:value-of select="cbc:ID/@schemeID"/> <xsl:text>: </xsl:text> <xsl:value-of select="cbc:ID"/> </td> </tr> </xsl:for-each> -->
										</tbody>
									</table>
									<!-- <hr/> -->
								</td>
								<!-- <td width="20%" align="center" valign="middle"> </td> -->
								<td width="28%" align="right" valign="top">
									<img style="width:220px;padding-top:33px;" align="middle" alt="ant logo" src="data:image/png;base64,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" />
								</td>
								<td>
								<div id="qrcode" style="position:relative; float:right; padding:10px;"/>
<div id="qrvalue" style="visibility: hidden; height: 15px;width: 15px; ; display:none"> 
{"vkntckn": "<xsl:value-of select="n1:Invoice/cac:AccountingSupplierParty/cac:Party/cac:PartyIdentification/cbc:ID[@schemeID = 'TCKN' or @schemeID = 'VKN']"/>",
"avkntckn": "<xsl:value-of select="n1:Invoice/cac:AccountingCustomerParty/cac:Party/cac:PartyIdentification/cbc:ID[@schemeID = 'TCKN' or @schemeID = 'VKN']"/><xsl:text> </xsl:text>",
"senaryo": "<xsl:value-of select="n1:Invoice/cbc:ProfileID"/>",
"tip": "<xsl:value-of select="n1:Invoice/cbc:InvoiceTypeCode"/>",
"tarih": "<xsl:value-of select="n1:Invoice/cbc:IssueDate"/>",
"no": "<xsl:value-of select="n1:Invoice/cbc:ID"/>",
"ettn": "<xsl:value-of select="n1:Invoice/cbc:UUID"/>",
"parabirimi": "<xsl:value-of select="n1:Invoice/cbc:DocumentCurrencyCode"/>",
"malhizmettoplam": "<xsl:value-of select="n1:Invoice/cac:LegalMonetaryTotal/cbc:LineExtensionAmount"/><xsl:for-each select="n1:Invoice/cac:TaxTotal/cac:TaxSubtotal[cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode = '0015']">"<xsl:text>,
"kdvmatrah</xsl:text>(<xsl:value-of select="cbc:Percent"/>)":"<xsl:value-of select="cbc:TaxableAmount"/>"</xsl:for-each><xsl:for-each select="n1:Invoice/cac:TaxTotal/cac:TaxSubtotal[cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode = '0015']"><xsl:text>,
"hesaplanankdv</xsl:text>(<xsl:value-of select="cbc:Percent"/>)":"<xsl:value-of select="cbc:TaxAmount"/>",</xsl:for-each>
"vergidahil":"<xsl:value-of select="n1:Invoice/cac:LegalMonetaryTotal/cbc:TaxInclusiveAmount"/>",
"odenecek":"<xsl:value-of select="n1:Invoice/cac:LegalMonetaryTotal/cbc:PayableAmount"/>"}</div>
									<script type="text/javascript">
										var qrcode = new QRCode(document.getElementById("qrcode"), {
											width : 200,
											height : 200,
											correctLevel : QRCode.CorrectLevel.H
										});

										function makeCode (msg) {		
											var elText = document.getElementById("text");
	
											qrcode.makeCode(msg);
										}

										makeCode(document.getElementById("qrvalue").innerHTML);
									</script>
								</td>
								<!-- <td width="40%"/> -->
							</tr>
						</tbody>
					</table>
					<table style="border-color:blue; " border="0" cellspacing="0px" width="800" cellpadding="0px">
						<tbody>
							<tr style="height:118px; " valign="top">
								<td width="40%" align="right" valign="bottom">
									<table id="customerPartyTable" align="left" border="0" height="50%">
										<tbody>
											<tr style="height:71px; ">
												<td>
													<hr />
													<table align="center" border="0">
														<tbody>
															<tr>
																<xsl:for-each select="n1:Invoice">
																	<xsl:for-each select="cac:AccountingCustomerParty">
																		<xsl:for-each select="cac:Party">
																			<td style="width:469px; " align="left">
																				<span style="font-weight:bold; ">
																					<xsl:text>SAYIN</xsl:text>
																				</span>
																			</td>
																		</xsl:for-each>
																	</xsl:for-each>
																</xsl:for-each>
															</tr>
															<tr>
																<xsl:for-each select="n1:Invoice">
																	<xsl:for-each select="cac:AccountingCustomerParty">
																		<xsl:for-each select="cac:Party">
																			<td style="width:469px; " align="left">
																				<xsl:if test="cac:PartyName">
																					<xsl:value-of select="cac:PartyName/cbc:Name" />
																					<br />
																				</xsl:if>
																				<xsl:for-each select="cac:Person">
																					<xsl:for-each select="cbc:Title">
																						<xsl:apply-templates />
																						<span>
																							<xsl:text> </xsl:text>
																						</span>
																					</xsl:for-each>
																					<xsl:for-each select="cbc:FirstName">
																						<xsl:apply-templates />
																						<span>
																							<xsl:text> </xsl:text>
																						</span>
																					</xsl:for-each>
																					<xsl:for-each select="cbc:MiddleName">
																						<xsl:apply-templates />
																						<span>
																							<xsl:text>  </xsl:text>
																						</span>
																					</xsl:for-each>
																					<xsl:for-each select="cbc:FamilyName">
																						<xsl:apply-templates />
																						<span>
																							<xsl:text> </xsl:text>
																						</span>
																					</xsl:for-each>
																					<xsl:for-each select="cbc:NameSuffix">
																						<xsl:apply-templates />
																					</xsl:for-each>
																				</xsl:for-each>
																			</td>
																		</xsl:for-each>
																	</xsl:for-each>
																</xsl:for-each>
															</tr>
															<tr>
																<xsl:for-each select="n1:Invoice">
																	<xsl:for-each select="cac:AccountingCustomerParty">
																		<xsl:for-each select="cac:Party">
																			<td style="width:469px; " align="left">
																				<xsl:for-each select="cac:PostalAddress">
																					<xsl:for-each select="//n1:Invoice/cbc:Note">
																						<xsl:if test="contains(.,'Delivery')">
																							<xsl:value-of select="substring-after(.,'Delivery; ')"/>
																						</xsl:if>
																					</xsl:for-each>
																					<xsl:for-each select="cbc:BuildingName">
																						<xsl:apply-templates />
																					</xsl:for-each>
																					<xsl:for-each select="cbc:BuildingNumber">
																						<span>
																							<xsl:text> No:</xsl:text>
																						</span>
																						<xsl:apply-templates />
																						<span>
																							<xsl:text> </xsl:text>
																						</span>
																					</xsl:for-each>
																				
																					<xsl:for-each select="cbc:Room">
																						<span>
																							<xsl:text>Kapı No:</xsl:text>
																						</span>
																						<xsl:apply-templates />
																						<span>
																							<xsl:text> </xsl:text>
																						</span>
																					</xsl:for-each>
																			
																					<xsl:for-each select="cbc:PostalZone">
																						<xsl:apply-templates />
																						<span>
																							<xsl:text> </xsl:text>
																						</span>
																					</xsl:for-each>
																					<xsl:for-each select="cbc:StreetName">
																						<xsl:apply-templates />
																						<span>
																							<xsl:text> </xsl:text>
																						</span>
																					</xsl:for-each>
																					<xsl:for-each select="cbc:CitySubdivisionName">
																						<xsl:apply-templates />
																						<span>
																							<xsl:text>/ </xsl:text>
																						</span>
																					</xsl:for-each>
																					
																					<xsl:for-each select="cbc:CityName">
																						<xsl:apply-templates />
																						<span>
																							<xsl:text> </xsl:text>
																						</span>
																					</xsl:for-each>
																				</xsl:for-each>
																			</td>
																		</xsl:for-each>
																	</xsl:for-each>
																</xsl:for-each>
															</tr>
															<xsl:for-each select="//n1:Invoice/cac:AccountingCustomerParty/cac:Party/cbc:WebsiteURI">
																<tr align="left">
																	<td>
																		<xsl:text>Web Sitesi: </xsl:text>
																		<xsl:value-of select="." />
																	</td>
																</tr>
															</xsl:for-each>
															<xsl:for-each select="//n1:Invoice/cac:AccountingCustomerParty/cac:Party/cac:Contact/cbc:ElectronicMail">
																<tr align="left">
																	<td>
																		<xsl:text>E-Posta: </xsl:text>
																		<xsl:value-of select="." />
																	</td>
																</tr>
															</xsl:for-each>
															<xsl:for-each select="n1:Invoice">
																<xsl:for-each select="cac:AccountingCustomerParty">
																	<xsl:for-each select="cac:Party">
																		<xsl:for-each select="cac:Contact">
																			<xsl:if test="cbc:Telephone or cbc:Telefax">
																				<tr align="left">
																					<td style="width:469px; " align="left">
																						<xsl:for-each select="cbc:Telephone">
																							<span>
																								<xsl:text>Tel: </xsl:text>
																							</span>
																							<xsl:apply-templates />
																						</xsl:for-each>
																						<xsl:for-each select="cbc:Telefax">
																							<span>
																								<xsl:text> Fax: </xsl:text>
																							</span>
																							<xsl:apply-templates />
																						</xsl:for-each>
																						<span>
																							<xsl:text> </xsl:text>
																						</span>
																					</td>
																				</tr>
																			</xsl:if>
																		</xsl:for-each>
																	</xsl:for-each>
																</xsl:for-each>
															</xsl:for-each>
															<xsl:if test="//n1:Invoice/cac:AccountingCustomerParty/cac:Party/cac:PartyTaxScheme/cac:TaxScheme/cbc:Name">
																<tr align="left">
																	<td>
																		<span>
																			<xsl:text>Vergi Dairesi: </xsl:text>
																			<xsl:value-of select="//n1:Invoice/cac:AccountingCustomerParty/cac:Party/cac:PartyTaxScheme/cac:TaxScheme/cbc:Name" />
																		</span>
																	</td>
																</tr>
															</xsl:if>
															<xsl:for-each select="//n1:Invoice/cac:AccountingCustomerParty/cac:Party/cac:PartyIdentification">
																<tr align="left">
																	<td>
																		<xsl:value-of select="cbc:ID/@schemeID" />
																		<xsl:text>: </xsl:text>
																		<xsl:value-of select="cbc:ID" />
																	</td>
																</tr>
															</xsl:for-each>
														</tbody>
													</table>
													<hr />
												</td>
											</tr>
										</tbody>
									</table>
									<br />
								</td>
								<td width="20%" align="center" valign="middle">
									<br />
									<br />
									<img style="width:91px;" align="middle" alt="E-Arşiv Logo" src="data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/4QBoRXhpZgAASUkqAAgAAAADABIBAwABAAAAAQAAADEBAgAQAAAAMgAAAGmHBAABAAAAQgAAAAAAAABTaG90d2VsbCAwLjIyLjAAAgACoAkAAQAAAKYBAAADoAkAAQAAAKYBAAAAAAAA/+EJ9Gh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8APD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iWE1QIENvcmUgNC40LjAtRXhpdjIiPiA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPiA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczpleGlmPSJodHRwOi8vbnMuYWRvYmUuY29tL2V4aWYvMS4wLyIgeG1sbnM6dGlmZj0iaHR0cDovL25zLmFkb2JlLmNvbS90aWZmLzEuMC8iIGV4aWY6UGl4ZWxYRGltZW5zaW9uPSI0MjIiIGV4aWY6UGl4ZWxZRGltZW5zaW9uPSI0MjIiIHRpZmY6SW1hZ2VXaWR0aD0iNDIyIiB0aWZmOkltYWdlSGVpZ2h0PSI0MjIiIHRpZmY6T3JpZW50YXRpb249IjEiLz4gPC9yZGY6UkRGPiA8L3g6eG1wbWV0YT4gICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICA8P3hwYWNrZXQgZW5kPSJ3Ij8+/9sAQwADAgIDAgIDAwMDBAMDBAUIBQUE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" />
									<h1 align="center">
										<span style="font-weight:bold; ">
											<xsl:text>e-Arşiv</xsl:text>
										</span>
									</h1>
									<img src='data:image/png;base64,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'/>
								</td>
								<td width="60%" align="center" valign="bottom" colspan="2">
									<!-- <table border="1" height="13" id="despatchTable"> -->
									<table height="13" id="despatchTable">
										<tbody>
											<!-- Uyumsoft irsaliye satirlarini otomatik ekliyor (select="n1:Invoice/cac:DespatchDocumentReference").
											     Satirlar Invoice baglaminda uretildigi icin eklenen blok hicbir sey basmaz. -->
											<xsl:for-each select="n1:Invoice">
											<tr>
												<td style="width:105px;" align="left">
													<span style="font-weight:bold; ">
														<xsl:text>Özelleştirme No:</xsl:text>
													</span>
												</td>
												<td style="width:110px;" align="left">
													<xsl:for-each select=".">
														<xsl:for-each select="cbc:CustomizationID">
															<xsl:apply-templates />
														</xsl:for-each>
													</xsl:for-each>
												</td>
											</tr>
											<tr style="height:13px; ">
												<td align="left">
													<span style="font-weight:bold; ">
														<xsl:text>Senaryo:</xsl:text>
													</span>
												</td>
												<td align="left">
													<xsl:for-each select=".">
														<xsl:for-each select="cbc:ProfileID">
															<xsl:apply-templates />
														</xsl:for-each>
													</xsl:for-each>
												</td>
											</tr>
											<tr style="height:13px; ">
												<td align="left">
													<span style="font-weight:bold; ">
														<xsl:text>Fatura Tipi:</xsl:text>
													</span>
												</td>
												<td align="left">
													<xsl:for-each select=".">
														<xsl:for-each select="cbc:InvoiceTypeCode">
															<xsl:apply-templates />
														</xsl:for-each>
													</xsl:for-each>
												</td>
											</tr>
											<tr style="height:13px; ">
												<td align="left">
													<span style="font-weight:bold; ">
														<xsl:text>Fatura No:</xsl:text>
													</span>
												</td>
												<td align="left">
													<xsl:for-each select=".">
														<xsl:for-each select="cbc:ID">
															<xsl:apply-templates />
														</xsl:for-each>
													</xsl:for-each>
												</td>
											</tr>
											<tr style="height:13px; ">
												<td align="left">
													<span style="font-weight:bold; ">
														<xsl:text>Fatura Tarihi:</xsl:text>
													</span>
												</td>
												<td align="left">
													<xsl:for-each select=".">
														<xsl:for-each select="cbc:IssueDate">
															<xsl:value-of select="substring(.,9,2)" />-<xsl:value-of select="substring(.,6,2)" />-<xsl:value-of select="substring(.,1,4)" />
														</xsl:for-each>
													</xsl:for-each>
												</td>
											</tr>
											</xsl:for-each>
										</tbody>
										<tbody>
											<xsl:for-each select="n1:Invoice">
											<xsl:if test="//n1:Invoice/cac:OrderReference">
												<tr style="height:13px">
													<td align="left">
														<span style="font-weight:bold; ">
															<xsl:text>Sipariş No:</xsl:text>
														</span>
													</td>
													<td align="left">
														<xsl:for-each select="cac:OrderReference">
															<xsl:for-each select="cbc:ID">
																<xsl:apply-templates />
															</xsl:for-each>
														</xsl:for-each>
													</td>
												</tr>
											</xsl:if>
											<xsl:if test="//n1:Invoice/cac:OrderReference/cbc:IssueDate">
												<tr style="height:13px">
													<td align="left">
														<span style="font-weight:bold; ">
															<xsl:text>Sipariş Tarihi:</xsl:text>
														</span>
													</td>
													<td align="left">
														<xsl:for-each select="cac:OrderReference">
															<xsl:for-each select="cbc:IssueDate">
																<xsl:value-of select="substring(.,9,2)" />-<xsl:value-of select="substring(.,6,2)" />-<xsl:value-of select="substring(.,1,4)" />
															</xsl:for-each>
														</xsl:for-each>
													</td>
												</tr>
											</xsl:if>
											</xsl:for-each>
										</tbody>
									</table>
								</td>
							</tr>
							<tr align="left">
								<table id="ettnTable">
									<tr style="height:13px;">
										<td align="left" valign="top">
											<span style="font-weight:bold; ">
												<xsl:text>ETTN:</xsl:text>
											</span>
										</td>
										<td align="left" width="240px">
											<xsl:for-each select="n1:Invoice">
												<xsl:for-each select="cbc:UUID">
													<xsl:apply-templates />
												</xsl:for-each>
											</xsl:for-each>
										</td>
									</tr>
								</table>
							</tr>
						</tbody>
					</table>
					<div id="lineTableAligner">
						<span>
							<xsl:text> </xsl:text>
						</span>
					</div>
					<table border="1" id="lineTable" width="800">
						<tbody>
							<tr id="lineTableTr">
								<td id="lineTableTd" style="width:3%">
									<span style="font-weight:bold; " align="center">
										<xsl:text>Sıra No</xsl:text>
									</span>
								</td>
								<td id="lineTableTd" align="center" style="width:10%">
									<span style="font-weight:bold; ">
										<xsl:text>Mal Hizmet Kodu</xsl:text>
									</span>
								</td>
								<td id="lineTableTd" style="width:20%" align="center">
									<span style="font-weight:bold; ">
										<xsl:text>Mal Hizmet</xsl:text>
									</span>
								</td>
								<td id="lineTableTd" style="width:7.4%" align="center">
									<span style="font-weight:bold;">
										<xsl:text>Miktar</xsl:text>
									</span>
								</td>
								<td id="lineTableTd" style="width:9%" align="center">
									<span style="font-weight:bold; ">
										<xsl:text>Birim Fiyat</xsl:text>
									</span>
								</td>
								<td id="lineTableTd" style="width:7%" align="center">
									<span style="font-weight:bold; ">
										<xsl:text>İskonto Oranı</xsl:text>
									</span>
								</td>
								<td id="lineTableTd" style="width:9%" align="center">
									<span style="font-weight:bold; ">
										<xsl:text>İskonto Tutarı</xsl:text>
									</span>
								</td>
								<td id="lineTableTd" style="width:7%" align="center">
									<span style="font-weight:bold; ">
										<xsl:text>KDV Oranı</xsl:text>
									</span>
								</td>
								<td id="lineTableTd" style="width:10%" align="center">
									<span style="font-weight:bold; ">
										<xsl:text>KDV Tutarı</xsl:text>
									</span>
								</td>
								<xsl:for-each select="//n1:Invoice/cac:InvoiceLine/cac:TaxTotal/cac:TaxSubtotal">
									<xsl:if test="./cbc:Percent=0">
										<xsl:if test="./cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode='0015' ">
											<td id="lineTableTd" style="width:10%" align="center">
												<span style="font-weight:bold; ">
													<xsl:text>KDV Muafiyet Sebebi</xsl:text>
												</span>
											</td>
										</xsl:if>
									</xsl:if>
								</xsl:for-each>
								<xsl:if test="//n1:Invoice/cac:InvoiceLine/cac:TaxTotal/cac:TaxSubtotal/cac:TaxCategory/cac:TaxScheme">
									<td id="lineTableTd" style="width:8%; " align="center">
										<span style="font-weight:bold; ">
											<xsl:text>Diğer Vergiler</xsl:text>
										</span>
									</td>
								</xsl:if>
								<td id="lineTableTd" style="width:10.6%" align="center">
									<span style="font-weight:bold; ">
										<xsl:text>Mal Hizmet Tutarı</xsl:text>
									</span>
								</td>
							</tr>
							<xsl:for-each select="//n1:Invoice/cac:InvoiceLine">
								<xsl:choose>
									<xsl:when test=".">
										<xsl:apply-templates select="." />
									</xsl:when>
									<xsl:otherwise>
										<xsl:apply-templates select="//n1:Invoice" />
									</xsl:otherwise>
								</xsl:choose>
							</xsl:for-each>
						</tbody>
					</table>
				</xsl:for-each>
				<br />
				<table id="budgetContainerTable" width="800px">
				<!-- <xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:LineExtensionAmount[(contains($parabirimilist, concat(' ', @currencyID, ' ')))]"> -->
					 <tr id="budgetContainerTr" align="right">
						<td id="budgetContainerDummyTd" />
						<td id="lineTableBudgetTd" align="right" width="200px">
							<span style="font-weight:bold; ">
								<xsl:text>Mal Hizmet Toplam Tutarı</xsl:text>
							</span>
						</td>
						<td id="lineTableBudgetTd" style="width:81px; " align="right">
							<span>
								<xsl:value-of select="format-number(//n1:Invoice/cac:LegalMonetaryTotal/cbc:LineExtensionAmount, '###.##0,00', 'european')" />
								<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:LineExtensionAmount/@currencyID">
									<xsl:text />
									<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:LineExtensionAmount[contains($parabirimilist, concat(' ', @currencyID, ' '))]">
										<xsl:text>TL</xsl:text>
									</xsl:if>
									<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:LineExtensionAmount[not(contains($parabirimilist, concat(' ', @currencyID, ' ')))]">
										<xsl:value-of select="//n1:Invoice/cac:LegalMonetaryTotal/cbc:LineExtensionAmount/@currencyID" />
									</xsl:if>
								</xsl:if>
							</span>
						</td>
					</tr>
					<tr id="budgetContainerTr" align="right">
						<td id="budgetContainerDummyTd" />
						<td id="lineTableBudgetTd" align="right" width="200px">
							<span style="font-weight:bold; ">
								<xsl:text>Toplam İskonto</xsl:text>
							</span>
						</td>
						<td id="lineTableBudgetTd" style="width:81px; " align="right">
							<span>
								<xsl:value-of select="format-number(//n1:Invoice/cac:LegalMonetaryTotal/cbc:AllowanceTotalAmount, '###.##0,00', 'european')" />
								<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:AllowanceTotalAmount/@currencyID">
									<xsl:text />
									<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:AllowanceTotalAmount[contains($parabirimilist, concat(' ', @currencyID, ' '))]">
										<xsl:text>TL</xsl:text>
									</xsl:if>
									<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:AllowanceTotalAmount[not(contains($parabirimilist, concat(' ', @currencyID, ' ')))]">
										<xsl:value-of select="//n1:Invoice/cac:LegalMonetaryTotal/cbc:AllowanceTotalAmount/@currencyID" />
									</xsl:if>
								</xsl:if>
							</span>
						</td>
					</tr>
					<tr id="budgetContainerTr" align="right">
						<td id="budgetContainerDummyTd" />
						<td id="lineTableBudgetTd" align="right" width="200px">
							<span style="font-weight:bold; ">
								<xsl:text>KDV Matrahı</xsl:text>
							</span>
						</td>
						<td id="lineTableBudgetTd" style="width:81px; " align="right">
							<span>
								<xsl:value-of select="format-number(//n1:Invoice/cac:LegalMonetaryTotal/cbc:TaxExclusiveAmount, '###.##0,00', 'european')" />
								<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:TaxExclusiveAmount/@currencyID">
									<xsl:text />
									<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:TaxExclusiveAmount[contains($parabirimilist, concat(' ', @currencyID, ' '))]">
										<xsl:text>TL</xsl:text>
									</xsl:if>
									<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:TaxExclusiveAmount[not(contains($parabirimilist, concat(' ', @currencyID, ' ')))]">
										<xsl:value-of select="//n1:Invoice/cac:LegalMonetaryTotal/cbc:TaxExclusiveAmount/@currencyID" />
									</xsl:if>
								</xsl:if>
							</span>
						</td>
					</tr>
					<xsl:for-each select="n1:Invoice/cac:TaxTotal/cac:TaxSubtotal">
						<tr id="budgetContainerTr" align="right">
							<td id="budgetContainerDummyTd" />
							<td id="lineTableBudgetTd" width="211px" align="right">
								<span style="font-weight:bold; ">
									<xsl:text>Hesaplanan </xsl:text>
									<xsl:value-of select="cac:TaxCategory/cac:TaxScheme/cbc:Name" />
									<xsl:text>(%</xsl:text>
									<xsl:value-of select="cbc:Percent" />
									<xsl:text>)</xsl:text>
								</span>
							</td>
							<td id="lineTableBudgetTd" style="width:82px; " align="right">
								<xsl:for-each select="cac:TaxCategory/cac:TaxScheme">
									<xsl:text />
									<xsl:value-of select="format-number(../../cbc:TaxAmount, '###.##0,00', 'european')" />
									<xsl:if test="../../cbc:TaxAmount/@currencyID">
										<xsl:text />
										<xsl:if test="../../cbc:TaxAmount[contains($parabirimilist, concat(' ', @currencyID, ' '))]">
											<xsl:text>TL</xsl:text>
										</xsl:if>
										<xsl:if test="../../cbc:TaxAmount[not(contains($parabirimilist, concat(' ', @currencyID, ' ')))]">
											<xsl:value-of select="../../cbc:TaxAmount/@currencyID" />
										</xsl:if>
									</xsl:if>
								</xsl:for-each>
							</td>
						</tr>
					</xsl:for-each>
					<xsl:for-each select="n1:Invoice/cac:WithholdingTaxTotal/cac:TaxSubtotal">
						<tr id="budgetContainerTr" align="right">
							<td id="budgetContainerDummyTd" />
							<td id="lineTableBudgetTd" width="211px" align="right">
								<span style="font-weight:bold; ">
									<xsl:text>Hesaplanan KDV TEVKİFAT</xsl:text>
									<!--<xsl:value-of select="cac:TaxCategory/cac:TaxScheme/cbc:Name"/>-->
									<xsl:text>(</xsl:text>
									<xsl:value-of select="cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode" />
									<xsl:text>)</xsl:text>
									<xsl:text> (%</xsl:text>
									<xsl:value-of select="cbc:Percent" />
									<xsl:text>)</xsl:text>
								</span>
							</td>
							<td id="lineTableBudgetTd" align="right=">
								<xsl:for-each select="cac:TaxCategory/cac:TaxScheme">
									<xsl:text />
									<xsl:value-of select="format-number(../../cbc:TaxAmount, '###.##0,00', 'european')" />
									<xsl:if test="../../cbc:TaxAmount/@currencyID">
										<!--<xsl:text> </xsl:text>
                    <xsl:if test="../../cbc:TaxAmount/@currencyID = 'TRY'">
                      <xsl:text>TL</xsl:text>
                    </xsl:if>
                    <xsl:if test="../../cbc:TaxAmount/@currencyID != 'TRY'">
                      <xsl:value-of select="../../cbc:TaxAmount/@currencyID"/>
                    </xsl:if>-->
										<xsl:text />
										<xsl:if test="../../cbc:TaxAmount[contains($parabirimilist, concat(' ', @currencyID, ' '))]">
											<xsl:text>TL</xsl:text>
										</xsl:if>
										<xsl:if test="../../cbc:TaxAmount[not(contains($parabirimilist, concat(' ', @currencyID, ' ')))]">
											<xsl:value-of select="../../cbc:TaxAmount/@currencyID" />
										</xsl:if>
									</xsl:if>
								</xsl:for-each>
							</td>
						</tr>
					</xsl:for-each>
					<xsl:if test ="//n1:Invoice/cbc:InvoiceTypeCode = 'TEVKIFAT'">
					<tr id="budgetContainerTr" align="right">
					<xsl:variable name="hesaplananKdv" select="//n1:Invoice/cac:TaxTotal/cac:TaxSubtotal/cbc:TaxAmount"/>	
	<xsl:variable name="teklifedilenKdv" select="//n1:Invoice/cac:WithholdingTaxTotal/cac:TaxSubtotal/cbc:TaxAmount"/>													
							<td id="budgetContainerDummyTd" />
							<td id="lineTableBudgetTd" width="211px" align="right">
								<span style="font-weight:bold; ">
									<xsl:text>Tahsil Edilen KDV</xsl:text>								
								</span>
							</td>
							<td id="lineTableBudgetTd" align="right=">
									<xsl:text />
									<xsl:value-of select="format-number($hesaplananKdv - $teklifedilenKdv, '###.##0,00', 'european')"/>
<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:TaxExclusiveAmount/@currencyID">
									<xsl:text />
									<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:TaxExclusiveAmount[contains($parabirimilist, concat(' ', @currencyID, ' '))]">
										<xsl:text>TL</xsl:text>
									</xsl:if>
									<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:TaxExclusiveAmount[not(contains($parabirimilist, concat(' ', @currencyID, ' ')))]">
										<xsl:value-of select="//n1:Invoice/cac:LegalMonetaryTotal/cbc:TaxExclusiveAmount/@currencyID" />
									</xsl:if>
								</xsl:if>
							</td>
						</tr>
						</xsl:if>
					<tr id="budgetContainerTr" align="right">
						<td id="budgetContainerDummyTd"/>
						<td id="lineTableBudgetTd" width="200px" align="right">
							<span style="font-weight:bold; ">
								<xsl:text>Vergiler Dahil Toplam Tutar</xsl:text>
							</span>
						</td>
						<td id="lineTableBudgetTd" style="width:82px; " align="right">
							<xsl:for-each select="n1:Invoice">
								<xsl:for-each select="cac:LegalMonetaryTotal">
									<xsl:for-each select="cbc:TaxInclusiveAmount">
										<xsl:value-of select="format-number(., '###.##0,00', 'european')"/>
										<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:TaxInclusiveAmount/@currencyID">
											<xsl:text/>
											<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:TaxInclusiveAmount[contains($parabirimilist, concat(' ', @currencyID, ' '))]">
												<xsl:text>TL</xsl:text>
											</xsl:if>
											<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:TaxInclusiveAmount[not(contains($parabirimilist, concat(' ', @currencyID, ' ')))]">
												<xsl:value-of select="//n1:Invoice/cac:LegalMonetaryTotal/cbc:TaxInclusiveAmount/@currencyID" />
											</xsl:if>
										</xsl:if>
									</xsl:for-each>
								</xsl:for-each>
							</xsl:for-each>
						</td>
					</tr>
					<tr id="budgetContainerTr" align="right">
						<td id="budgetContainerDummyTd" />
						<td id="lineTableBudgetTd" width="200px" align="right">
							<span style="font-weight:bold; ">
								<xsl:text>Ödenecek Tutar</xsl:text>
							</span>
						</td>
						<td id="lineTableBudgetTd" style="width:82px; " align="right">
							<xsl:for-each select="n1:Invoice">
								<xsl:for-each select="cac:LegalMonetaryTotal">
									<xsl:for-each select="cbc:PayableAmount">
										<xsl:value-of select="format-number(., '###.##0,00', 'european')" />
										<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:PayableAmount/@currencyID">
											<xsl:text />
											<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:PayableAmount[contains($parabirimilist, concat(' ', @currencyID, ' '))]">
												<xsl:text>TL</xsl:text>
											</xsl:if>
											<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:PayableAmount[not(contains($parabirimilist, concat(' ', @currencyID, ' ')))]">
												<xsl:value-of select="//n1:Invoice/cac:LegalMonetaryTotal/cbc:PayableAmount/@currencyID" />
											</xsl:if>
										</xsl:if>
									</xsl:for-each>
								</xsl:for-each>
							</xsl:for-each>
						</td>
					</tr>
					<!-- </xsl:if> -->
					
				</table>
				<br />
				<table id="notesTable" width="800" align="left" height="100">
					<tbody>
						<tr align="left">
							<td id="notesTableTd" colspan="4">
								<xsl:if test="//n1:Invoice/cbc:Note">
									<b> Not: </b>
									<xsl:for-each select="//n1:Invoice/cbc:Note">   <xsl:value-of select="." />
										<xsl:text> </xsl:text>
										<br />
									</xsl:for-each>
								</xsl:if>
								<xsl:if test="//n1:Invoice/cac:PaymentMeans/cbc:InstructionNote">
									<b> Ödeme Notu: </b>
									<xsl:value-of select="//n1:Invoice/cac:PaymentMeans/cbc:InstructionNote" />
									<br />
								</xsl:if>
								<xsl:if test="//n1:Invoice/cac:PaymentMeans/cac:PayeeFinancialAccount/cbc:PaymentNote">
									<b> Hesap Açıklaması: </b>
									<xsl:value-of select="//n1:Invoice/cac:PaymentMeans/cac:PayeeFinancialAccount/cbc:PaymentNote" />
									<br />
								</xsl:if>
								<xsl:if test="//n1:Invoice/cac:PaymentTerms/cbc:Note">
									<b> Ödeme Koşulu: </b>
									<xsl:value-of select="//n1:Invoice/cac:PaymentTerms/cbc:Note" />
									<br />
								</xsl:if>
								<xsl:if test="//n1:Invoice/cac:LegalMonetaryTotal/cbc:LineExtensionAmount/@currencyID != 'TRY'">
								  <b> <br></br>
    <br></br>
    ÖDEME NOTU / PAYMENT NOTE: " Döviz cinsinden veya Dövize endeksli olarak düzenlenmiş Fatura Ödemelerinde Döviz/TL   <br></br>  
    çeviriminde fiili ödeme günündeki QNB Finansbank Satış Kuru baz alınacak olup oluşacak Kur farkları taraflarca faturalandırılacaktır.<br></br>
		
</b><br></br>Not: Muhasebe Kayıt Mat: <xsl:value-of select="format-number(//n1:Invoice/cac:LegalMonetaryTotal/cbc:LineExtensionAmount * //n1:Invoice/cac:PricingExchangeRate/cbc:CalculationRate, '###.##0,0000', 'european')"/>
									<xsl:text> TL</xsl:text> KDV:  	<xsl:value-of select="format-number(//n1:Invoice/cac:TaxTotal/cbc:TaxAmount * //n1:Invoice/cac:PricingExchangeRate/cbc:CalculationRate, '###.##0,0000', 'european')"/>
								<xsl:text> TL</xsl:text> Genel Toplam: <xsl:value-of select="format-number(//n1:Invoice/cac:LegalMonetaryTotal/cbc:PayableAmount * //n1:Invoice/cac:PricingExchangeRate/cbc:CalculationRate, '###.##0,0000', 'european')"/>
								<xsl:text> TL</xsl:text>
		</xsl:if><br></br>
		<b>Lütfen ödemelerinizi yaparken açıklama kısmına fatura numaramızı belirtiniz.</b>
							</td>
						</tr>
						<!-- <tr align="left" valign="top"> -->
							<!-- <td id="notesTableTd" colspan="4"> -->
								<hr/>
								<!-- <br /> -->
								<!-- <u> -->
									<!-- <b> -->
										<!-- <xsl:text>Banka Hesaplarımız</xsl:text> -->
									<!-- </b> -->
								<!-- </u> -->
							<!-- </td> -->
						<!-- </tr> -->
				<tr align="left" valign="top">
							<td id="notesTableTd" colspan="4">
								<!-- <hr/> -->
								<br/>
								<u>
									<b>
										<xsl:text>Banka Hesaplarımız</xsl:text>
									</b>
								</u>
							</td>
						</tr>
						<table id="tablebank" width="800" style="font-size:9px;  ">
    <thead>
        <tr align="left">
            <th>Banka</th>
            <th>TL IBAN</th>
            <th>EURO IBAN</th>
            <th>USD IBAN</th>
            <th>GBP IBAN</th>
        </tr>
    </thead>
    <tbody>
        <!-- QNB BANK -->
        <tr>
            <td>QNB BANK</td>
            <td>TR63 0011 1000 0000 0001 9644 00</td>
            <td>TR11 0011 1000 0000 0005 5943 97</td>
            <td>TR38 0011 1000 0000 0005 5943 96</td>
            <td>TR64 0011 1000 0000 0011 7518 76</td>
        </tr>

        <!-- YKB -->
        <tr>
            <td>YKB</td>
            <td>TR82 0006 7010 0000 0070 6144 28</td>
            <td>TR82 0006 7010 0000 0070 6821 34</td>
            <td>TR50 0006 7010 0000 0070 6821 28</td>
            <td>TR88 0006 7010 0000 0052 1591 97</td>
        </tr>

        <!-- GARANTİ BANKASI -->
        <tr>
            <td>GARANTİ BANKASI</td>
            <td>TR91 0006 2000 3410 0006 2959 76</td>
            <td>TR74 0006 2000 3410 0009 0906 58</td>
            <td>TR82 0006 2000 3410 0009 0844 03</td>
            <td>TR52 0006 2000 3410 0009 0788 32</td>
        </tr>

        <!-- HALKBANK -->
        <tr>
            <td>HALKBANK</td>
            <td>TR09 0001 2001 5930 0010 1008 27</td>
            <td>TR87 0001 2001 5930 0058 1001 64</td>
            <td>TR87 0001 2001 5930 0053 1002 02</td>
            <td>TR62 0001 2001 5930 0055 1000 16</td>
        </tr>

        <!-- ZİRAAT BANKASI -->
        <tr>
            <td>ZİRAAT BANKASI</td>
            <td>TR02 0001 0022 4802 8707 9750 50</td>
            <td>TR18 0001 0022 4802 8707 9750 53</td>
            <td>TR45 0001 0022 4802 8707 9750 52</td>
            <td>TR88 0001 0022 4802 8707 9750 54</td>
        </tr>
    </tbody>
</table>
						<tr align="left" valign="top">
							<td id="notesTableTd" colspan="4">
								<!-- <hr/> -->
								<br />
								<xsl:text>İşbu faturaya 7 gün içinde itiraz edilmediği takdirde kabul edilmiş sayılır.</xsl:text>
							</td>
						</tr>
					</tbody>
				</table>
			</body>
		</html>
	</xsl:template>
	<xsl:template match="dateFormatter">
		<xsl:value-of select="substring(.,9,2)" />-<xsl:value-of select="substring(.,6,2)" />-<xsl:value-of select="substring(.,1,4)" />
	</xsl:template>
	<xsl:template match="//n1:Invoice/cac:InvoiceLine">
		<tr id="lineTableTr">
			<td id="lineTableTd">
				<span>
					<xsl:text> </xsl:text>
					<xsl:value-of select="./cbc:ID" />
				</span>
			</td>
			<td id="lineTableTd">
				<span>
					<xsl:text> </xsl:text>
					<xsl:value-of select="./cac:Item/cac:SellersItemIdentification/cbc:ID" />
					<xsl:text> </xsl:text>
				</span>
			</td>
			<td id="lineTableTd">
				<span>
					<xsl:text> </xsl:text>
					<xsl:value-of select="./cac:Item/cbc:Name" />
				</span>
			</td>
			<td id="lineTableTd" align="right">
				<span>
					<xsl:text> </xsl:text>
					<xsl:value-of select="format-number(./cbc:InvoicedQuantity, '###.###,##', 'european')" />
					<xsl:text> </xsl:text>
					<xsl:if test="./cbc:InvoicedQuantity/@unitCode">
						<xsl:for-each select="./cbc:InvoicedQuantity">
							<xsl:text />
							<xsl:choose>
								<xsl:when test="@unitCode  = '26'">
									<span>
										<xsl:text>Ton</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'BX'">
									<span>
										<xsl:text>Kutu</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'LTR'">
									<span>
										<xsl:text>LT</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'NIU' or @unitCode  = 'C62'">
									<span>
										<xsl:text>Adet</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'SET'">
									<span>
										<xsl:text>SET</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'KGM'">
									<span>
										<xsl:text>KG</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'KJO'">
									<span>
										<xsl:text>kJ</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'GRM'">
									<span>
										<xsl:text>G</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'MGM'">
									<span>
										<xsl:text>MG</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'NT'">
									<span>
										<xsl:text>Net Ton</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'GT'">
									<span>
										<xsl:text>GT</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'MTR'">
									<span>
										<xsl:text>M</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'MMT'">
									<span>
										<xsl:text>MM</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'KTM'">
									<span>
										<xsl:text>KM</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'MLT'">
									<span>
										<xsl:text>ML</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'MMQ'">
									<span>
										<xsl:text>MM3</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'CLT'">
									<span>
										<xsl:text>CL</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'CMK'">
									<span>
										<xsl:text>CM2</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'CMQ'">
									<span>
										<xsl:text>CM3</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'CMT'">
									<span>
										<xsl:text>CM</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'MTK'">
									<span>
										<xsl:text>M2</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'MTQ'">
									<span>
										<xsl:text>M3</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'DAY'">
									<span>
										<xsl:text> Gün</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'MON'">
									<span>
										<xsl:text> Ay</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'PA'">
									<span>
										<xsl:text> Paket</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'KWH'">
									<span>
										<xsl:text> KWH</xsl:text>
									</span>
								</xsl:when>
								<xsl:when test="@unitCode  = 'BO'">
                  <span>
                    <xsl:text>Şişe</xsl:text>
                  </span>
                </xsl:when>
				 <xsl:when test="@unitCode  = 'TU'">
                  <span>
                    <xsl:text>Tüp</xsl:text>
                  </span>
                </xsl:when>
				 
							</xsl:choose>
						</xsl:for-each>
					</xsl:if>
				</span>
			</td>
			<td id="lineTableTd" align="right">
				<span>
					<xsl:text>&#160;</xsl:text>
					<xsl:value-of select="format-number(./cac:Price/cbc:PriceAmount, '###.##0,00###', 'european')"/>
					<xsl:if test="./cac:Price/cbc:PriceAmount/@currencyID">
						<xsl:text/>
						<xsl:if test="./cac:Price/cbc:PriceAmount[contains($parabirimilist, concat(' ', @currencyID, ' '))] ">
							<xsl:text>TL</xsl:text>
						</xsl:if>
						<xsl:if test="./cac:Price/cbc:PriceAmount[not(contains($parabirimilist, concat(' ', @currencyID, ' ')))]">
							<xsl:value-of select="./cac:Price/cbc:PriceAmount/@currencyID"/>
						</xsl:if>
					</xsl:if>
				</span>
			</td>
			<td id="lineTableTd" align="right">
				<span>
					<xsl:text> </xsl:text>
					<xsl:if test="./cac:AllowanceCharge/cbc:MultiplierFactorNumeric">
						<xsl:text> %</xsl:text>
						<xsl:value-of select="format-number(./cac:AllowanceCharge/cbc:MultiplierFactorNumeric * 100, '###.##0,00', 'european')" />
					</xsl:if>
				</span>
			</td>
			<td id="lineTableTd" align="right">
				<span>
					<xsl:text> </xsl:text>
					<xsl:if test="./cac:AllowanceCharge">
						<!--<xsl:if test="./cac:AllowanceCharge/cbc:ChargeIndicator = true() ">+ </xsl:if> <xsl:if test="./cac:AllowanceCharge/cbc:ChargeIndicator = false() ">- </xsl:if>-->
						<xsl:if
                     test="./cac:AllowanceCharge/cbc:Amount[not(contains($parabirimilist, concat(' ', @currencyID, ' ')))] and //n1:Invoice/cac:PricingExchangeRate/cbc:CalculationRate > 1">
					 <xsl:value-of select="format-number(./cac:AllowanceCharge/cbc:Amount * //n1:Invoice/cac:PricingExchangeRate/cbc:CalculationRate, '###.##0,00', 'european')" />
					 </xsl:if>
					 <xsl:if test="./cac:Price/cbc:PriceAmount[(contains($parabirimilist, concat(' ', @currencyID, ' ')))]">
					 <xsl:value-of select="format-number(./cac:AllowanceCharge/cbc:Amount, '###.##0,00', 'european')" />
					 </xsl:if>
					</xsl:if>
					 <xsl:text></xsl:text>
					
					<xsl:if test="./cac:AllowanceCharge/cbc:Amount/@currencyID">
						<xsl:text />
						<xsl:if test="./cac:AllowanceCharge/cbc:Amount[contains($parabirimilist, concat(' ', @currencyID, ' '))]">
							<xsl:text> TL</xsl:text>
						</xsl:if>
						<xsl:if test="./cac:AllowanceCharge/cbc:Amount[not(contains($parabirimilist, concat(' ', @currencyID, ' ')))]">
							<xsl:value-of select="./cac:AllowanceCharge/cbc:Amount/@currencyID" />
						</xsl:if>
					</xsl:if>
				</span>
			</td>
			<td id="lineTableTd" align="right">
				<span>
					<xsl:text> </xsl:text>
					<xsl:for-each select="./cac:TaxTotal/cac:TaxSubtotal/cac:TaxCategory/cac:TaxScheme">
						<xsl:if test="cbc:TaxTypeCode='0015' ">
							<xsl:text />
							<xsl:if test="../../cbc:Percent">
								<xsl:text> %</xsl:text>
								<xsl:value-of select="format-number(../../cbc:Percent, '###.##0,00', 'european')" />
							</xsl:if>
						</xsl:if>
					</xsl:for-each>
				</span>
			</td>
			<td id="lineTableTd" align="right">
				<span>
					<xsl:text>&#160;</xsl:text>
					<xsl:for-each select="./cac:TaxTotal/cac:TaxSubtotal/cac:TaxCategory/cac:TaxScheme">
						<xsl:if test="cbc:TaxTypeCode='0015' ">
							<xsl:text/>
							<xsl:value-of select="format-number(../../cbc:TaxAmount, '###.##0,00', 'european')"/>
							<xsl:if test="../../cbc:TaxAmount/@currencyID">
								<xsl:text/>
								<xsl:if test="../../cbc:TaxAmount[contains($parabirimilist, concat(' ', @currencyID, ' '))]">
									<xsl:text>TL</xsl:text>
								</xsl:if>
								<xsl:if test="../../cbc:TaxAmount[not(contains($parabirimilist, concat(' ', @currencyID, ' ')))]">
									<xsl:value-of select="../../cbc:TaxAmount/@currencyID"/>
								</xsl:if>
							</xsl:if>
						</xsl:if>
					</xsl:for-each>
				</span>
			</td>
			<xsl:variable name="kdv0" select="./cac:TaxTotal/cac:TaxSubtotal/cac:TaxCategory/cbc:TaxExemptionReason" />
			<xsl:for-each select="//n1:Invoice/cac:InvoiceLine/cac:TaxTotal/cac:TaxSubtotal">
				<xsl:if test="./cbc:Percent=0">
					<xsl:if test="./cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode='0015' ">
						<td id="lineTableTd" style="width:10%" align="center">
							<xsl:value-of select="$kdv0" />
						</td>
					</xsl:if>
				</xsl:if>
			</xsl:for-each>
			<xsl:if test="./cac:TaxTotal/cac:TaxSubtotal/cac:TaxCategory/cac:TaxScheme">
				<td id="lineTableTd" style="font-size: xx-small" align="right">
					<span>
						<xsl:text> </xsl:text>
						<xsl:for-each select="./cac:TaxTotal/cac:TaxSubtotal/cac:TaxCategory/cac:TaxScheme">
							<xsl:if test="cbc:TaxTypeCode!='0015' ">
								<xsl:text />
								<xsl:value-of select="cbc:Name" />
								<xsl:if test="../../cbc:Percent">
									<xsl:text> (%</xsl:text>
									<xsl:value-of select="format-number(../../cbc:Percent, '###.##0,00', 'european')" />
									<xsl:text>)=</xsl:text>
								</xsl:if>
								<xsl:value-of select="format-number(../../cbc:TaxAmount, '###.##0,00', 'european')" />
								<xsl:if test="../../cbc:TaxAmount/@currencyID">
									<xsl:text />
									<xsl:if test="../../cbc:TaxAmount[contains($parabirimilist, concat(' ', @currencyID, ' '))]">
										<xsl:text>TL</xsl:text>
									</xsl:if>
									<xsl:if test="../../cbc:TaxAmount[not(contains($parabirimilist, concat(' ', @currencyID, ' ')))]">
										<xsl:value-of select="../../cbc:TaxAmount/@currencyID" />
									</xsl:if>
								</xsl:if>
							</xsl:if>
						</xsl:for-each>
					</span>
				</td>
			</xsl:if>
			<td id="lineTableTd" align="right">
				<span>
					<xsl:text>&#160;</xsl:text>
					<xsl:value-of select="format-number(./cbc:LineExtensionAmount, '###.##0,00', 'european')"/>
					<xsl:if test="./cbc:LineExtensionAmount/@currencyID">
						<xsl:text/>
						<xsl:if test="./cbc:LineExtensionAmount[contains($parabirimilist, concat(' ', @currencyID, ' '))] ">
							<xsl:text>TL</xsl:text>
						</xsl:if>
						<xsl:if test="./cbc:LineExtensionAmount[not(contains($parabirimilist, concat(' ', @currencyID, ' ')))] ">
							<xsl:value-of select="./cbc:LineExtensionAmount/@currencyID"/>
						</xsl:if>
					</xsl:if>
				</span>
			</td>
		</tr>
	</xsl:template>
	<xsl:template match="//n1:Invoice">
		<tr id="lineTableTr">
			<td id="lineTableTd">
				<span>
					<xsl:text> </xsl:text>
				</span>
			</td>
			<td id="lineTableTd">
				<span>
					<xsl:text> </xsl:text>
				</span>
			</td>
			<td id="lineTableTd" align="right">
				<span>
					<xsl:text> </xsl:text>
				</span>
			</td>
			<td id="lineTableTd" align="right">
				<span>
					<xsl:text> </xsl:text>
				</span>
			</td>
			<td id="lineTableTd" align="right">
				<span>
					<xsl:text> </xsl:text>
				</span>
			</td>
			<td id="lineTableTd" align="right">
				<span>
					<xsl:text> </xsl:text>
				</span>
			</td>
			<td id="lineTableTd" align="right">
				<span>
					<xsl:text> </xsl:text>
				</span>
			</td>
			<td id="lineTableTd" align="right">
				<span>
					<xsl:text> </xsl:text>
				</span>
			</td>
			<td id="lineTableTd" align="right">
				<span>
					<xsl:text> </xsl:text>
				</span>
			</td>
			<td id="lineTableTd" align="right">
				<span>
					<xsl:text> </xsl:text>
				</span>
			</td>
		</tr>
	</xsl:template>
</xsl:stylesheet>