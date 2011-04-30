//PostScreenGlow: Used example from Rocker Commander for this and did some heavy modifications to it
//Removed a lot of the radial and shader 20 stuff as we want to keep shader11 as the base requirement

const float DownsampleMultiplicator =0.25f;
const float4 ClearColor : DIFFUSE = { 0.0f, 0.0f, 0.0f, 1.0f};
const float ClearDepth = 1.0f;

float Script : STANDARDSGLOBAL
<
	//string UIWidget = "none";
	string ScriptClass = "scene";
	string ScriptOrder = "postprocess";
	string ScriptOutput = "color";

	// We just call a script in the main technique.
	string Script = "Technique=ScreenGlow;"; 
> = 0.5;

float GlowIntensity <
    string UIName = "Glow intensity";
    string UIWidget = "slider";
    float UIMin = 0.0f;
    float UIMax = 1.0f;
    float UIStep = 0.1f;
> = 0.5f;//0.7f

float HighlightIntensity <
    string UIName = "Highlight intensity";
    string UIWidget = "slider";
    float UIMin = 0.0f;
    float UIMax = 1.0f;
    float UIStep = 0.1f;
> = 0.0f;//0.4f

// Render-to-Texture stuff
float2 windowSize : VIEWPORTPIXELSIZE;
const float downsampleScale = 0.25;

texture sceneMap : RENDERCOLORTARGET
< 
    float2 ViewportRatio = { 1.0, 1.0 };
    int MIPLEVELS = 1;
    string format = "X8R8G8B8";
>;
sampler sceneMapSampler = sampler_state 
{
    texture = <sceneMap>;
    AddressU  = CLAMP;
    AddressV  = CLAMP;
    AddressW  = CLAMP;
    MIPFILTER = NONE;
    MINFILTER = LINEAR;
    MAGFILTER = LINEAR;
};

texture downsampleMap : RENDERCOLORTARGET
< 
    float2 ViewportRatio = { 0.25, 0.25 };
    int MIPLEVELS = 1;
    string format = "A8R8G8B8";
>;
sampler downsampleMapSampler = sampler_state 
{
    texture = <downsampleMap>;
    AddressU  = CLAMP;        
    AddressV  = CLAMP;
    AddressW  = CLAMP;
    MIPFILTER = NONE;
    MINFILTER = LINEAR;
    MAGFILTER = LINEAR;
};

texture blurMap1 : RENDERCOLORTARGET
< 
    float2 ViewportRatio = { 0.25, 0.25 };
    int MIPLEVELS = 1;
    string format = "A8R8G8B8";
>;
sampler blurMap1Sampler = sampler_state 
{
    texture = <blurMap1>;
    AddressU  = CLAMP;        
    AddressV  = CLAMP;
    AddressW  = CLAMP;
    MIPFILTER = NONE;
    MINFILTER = LINEAR;
    MAGFILTER = LINEAR;
};

texture blurMap2 : RENDERCOLORTARGET
< 
    float2 ViewportRatio = { 0.25, 0.25 };
    int MIPLEVELS = 1;
    string format = "A8R8G8B8";
>;
sampler blurMap2Sampler = sampler_state 
{
    texture = <blurMap2>;
    AddressU  = CLAMP;        
    AddressV  = CLAMP;
    AddressW  = CLAMP;
    MIPFILTER = NONE;
    MINFILTER = LINEAR;
    MAGFILTER = LINEAR;
};

 
// Returns luminance value of col to convert color to grayscale
float Luminance(float3 col)
{
	return dot(col, float3(0.3, 0.59, 0.11));
} // Luminance(.)

struct VB_OutputPosTexCoord
{
   	float4 pos      : POSITION;
    float2 texCoord : TEXCOORD0;
};

struct VB_OutputPos2TexCoords
{
   	float4 pos         : POSITION;
    float2 texCoord[2] : TEXCOORD0;
};

struct VB_OutputPos3TexCoords
{
   	float4 pos         : POSITION;
    float2 texCoord[3] : TEXCOORD0;
};

struct VB_OutputPos4TexCoords
{
   	float4 pos         : POSITION;
    float2 texCoord[4] : TEXCOORD0;
};

float4 PS_Display(
	VB_OutputPosTexCoord In,
	uniform sampler2D tex) : COLOR
{   
	float4 outputColor = tex2D(tex, In.texCoord);
	// Display color
	return outputColor;
	//return float4(1, 0, 0, 1);
} // PS_Display(..)

float4 PS_DisplayAlpha(
	VB_OutputPosTexCoord In,
	uniform sampler2D tex) : COLOR
{
	float4 outputColor = tex2D(tex, In.texCoord);
	// Just display alpha
	return float4(outputColor.a, outputColor.a, outputColor.a, 0.0f);
} // PS_DisplayAlpha(..)

// Generate texture coordinates to only 2 sample neighbours (can't do more in ps)
VB_OutputPos2TexCoords VS_DownSample11(
	float4 pos      : POSITION,
	float2 texCoord : TEXCOORD0)
{
	VB_OutputPos2TexCoords Out = (VB_OutputPos2TexCoords)0;
	float2 texelSize = DownsampleMultiplicator /
		(windowSize * downsampleScale);
	float2 s = texCoord;
	Out.pos = pos;
#if 1
	Out.texCoord[0] = s - float2(-1, -1)*texelSize;
	Out.texCoord[1] = s - float2(+1, +1)*texelSize;
#else
	Out.texCoord[0] = s - float2(0, 0)*texelSize;
	Out.texCoord[1] = s - float2(+2, +2)*texelSize;
#endif
	return Out;
} // VS_DownSample11(..)

float4 PS_DownSample11(
	VB_OutputPos2TexCoords In,
	uniform sampler2D tex) : COLOR
{
	float4 c;

	// sub sampling (can't do more in ps_1_1)
	c = tex2D(tex, In.texCoord[0])/2;
	c += tex2D(tex, In.texCoord[1])/2;
	//c += tex2D(tex, In.texCoord[2])/4;
	//c += tex2D(tex, In.texCoord[3])/4;

	// store hilights in alpha, can't use smoothstep version!
#if 0 // texture lookup doesn't work for ps_1_1 without tex coord from vertex shader :(
	c.a = Highlights(c.rgb);
#else
	// Fake it with highly optimized version using 80% as treshold.
	float l = Luminance(c.rgb);
	float treshold = 0.75f;
	if (l < treshold)
		c.a = 0;
	else
	{
		l = l-treshold;
		l = l+l+l+l; // bring 0..0.25 back to 0..1
		c.a = l;
		//c.a = (l-treshold)+(l-treshold)+(l-treshold)+(l-treshold);//+(l-treshold)+(l-treshold);
		//+(l-0.8f)+(l-0.8f)+(l-0.8f);//+(l-0.8f);//*4;
		//l = l * l * l * l;
		 //saturate(l*2.0f-0.5f);//*4.0f-3.0f;
			//(l-0.8f)*5;
			//l 
		//we don't have any instructions left for this stuff:
		//l = l * l * (3-2 * l);
		//c.a = l;
	} // else
	//*/
	//c.a = l;
#endif

	return c;
	//return float4(l, l, l, c.g);	
	//return float4(c.r, c.g, c.b, 1);//l);
	//return float4(1, 0, 0, 1);
} // PS_DownSample11(..)


VB_OutputPos4TexCoords VS_SimpleBlur(
	uniform float2 direction,
	float4 pos      : POSITION, 
	float2 texCoord : TEXCOORD0)
{
	VB_OutputPos4TexCoords Out = (VB_OutputPos4TexCoords)0;
	Out.pos = pos;
	float2 texelSize = 1.0f / windowSize;

	Out.texCoord[0] = texCoord + texelSize*(float2(2.0f, 2.0f)+direction*(-3.0f));
	Out.texCoord[1] = texCoord + texelSize*(float2(2.0f, 2.0f)+direction*(-1.0f));
	Out.texCoord[2] = texCoord + texelSize*(float2(2.0f, 2.0f)+direction*(+1.0f));
	Out.texCoord[3] = texCoord + texelSize*(float2(2.0f, 2.0f)+direction*(+3.0f));
	
	return Out;
} // VS_SimpleBlur(..)

float4 PS_SimpleBlur(
	VB_OutputPos4TexCoords In,
	uniform sampler2D tex) : COLOR
{
	float4 OutputColor = 0;
	float fMult = 0.25f;
	OutputColor += tex2D(tex, In.texCoord[0])*fMult;
	OutputColor += tex2D(tex, In.texCoord[1])*fMult;
	OutputColor += tex2D(tex, In.texCoord[2])*fMult;
	OutputColor += tex2D(tex, In.texCoord[3])*fMult;
	return OutputColor;///4;
} // PS_SimpleBlur(..)


VB_OutputPos2TexCoords VS_ScreenQuad(
	float4 pos      : POSITION, 
	float2 texCoord : TEXCOORD0)
{
	VB_OutputPos2TexCoords Out;
	float2 texelSize = 1.0 /
		(windowSize * downsampleScale);
	Out.pos = pos;
	// Don't use bilinear filtering
	Out.texCoord[0] = texCoord + texelSize*0.5;
	Out.texCoord[1] = texCoord + texelSize*0.5;
	return Out;
} // VS_ScreenQuad(..)

VB_OutputPos3TexCoords VS_ScreenQuadSampleUp(
	float4 pos      : POSITION, 
	float2 texCoord : TEXCOORD0)
{
	VB_OutputPos3TexCoords Out;
	float2 texelSize = 1.0 / windowSize;
	Out.pos = pos;
	// Don't use bilinear filtering
	Out.texCoord[0] = texCoord + texelSize*0.5f;
	Out.texCoord[1] = texCoord + texelSize*0.5f/downsampleScale;
	Out.texCoord[2] = texCoord + (1.0/128.0f)*0.5f;
	return Out;
} // VS_ScreenQuadSampleUp(..)

float4 PS_ComposeFinalImage(
	VB_OutputPos3TexCoords In,
	uniform sampler2D sceneSampler,
	uniform sampler2D blurredSceneSampler) : COLOR
{
	float4 orig = tex2D(sceneSampler, In.texCoord[0]);
	float4 blur = tex2D(blurredSceneSampler, In.texCoord[1]);
	float4 ret = 0.75f * orig + GlowIntensity * blur;// + (HighlightIntensity * blur.g);
	return ret;
} // PS_ComposeFinalImage(...)

// Bloom technique for ps_1_1 (not that powerful, but looks still gooood)
technique ScreenGlow
<
	// Script stuff is just for FX Composer
	string Script =
		"ClearSetDepth=ClearDepth;"
		"RenderColorTarget=sceneMap;"
		//never used anyway: "RenderDepthStencilTarget=DepthMap;"
		"ClearSetColor=ClearColor;"
		"ClearSetDepth=ClearDepth;"
		"Clear=Color;"
		"Clear=Depth;"
		"ScriptSignature=color;"
		"ScriptExternal=;"
		"Pass=DownSample;"
		"Pass=GlowBlur1;"
		"Pass=GlowBlur2;"
		"Pass=ComposeFinalScene;";
>
{
	// Sample full render area down to (1/4, 1/4) of its size!
	pass DownSample
	<
		string Script =
			"RenderColorTarget0=downsampleMap;"
			"ClearSetColor=ClearColor;"
			"Clear=Color;"
			"Draw=Buffer;";
	>
	{
		//cullmode = none;
		//ZEnable = false;
		//ZWriteEnable = false;
		VertexShader = compile vs_1_1 VS_DownSample11();
		PixelShader  = compile ps_1_1 PS_DownSample11(sceneMapSampler);
	} // pass DownSample

	// Blur everything to make the glow effect.
	pass GlowBlur1
	<
		string Script =
			"RenderColorTarget0=blurMap1;"
			"ClearSetColor=ClearColor;"
			"Clear=Color;"
			"Draw=Buffer;";
	>
	{
		//cullmode = none;
		//ZEnable = false;
		//ZWriteEnable = false;
		VertexShader = compile vs_1_1 VS_SimpleBlur(float2(2, 0));
		PixelShader  = compile ps_1_1 PS_SimpleBlur(downsampleMapSampler);
	} // pass GlowBlur1

	pass GlowBlur2
	<
		string Script =
			"RenderColorTarget0=blurMap2;"
			"ClearSetColor=ClearColor;"
			"Clear=Color;"
			"Draw=Buffer;";
	>
	{
		//cullmode = none;
		//ZEnable = false;
		//ZWriteEnable = false;
		VertexShader = compile vs_1_1 VS_SimpleBlur(float2(0, 2));
		PixelShader  = compile ps_1_1 PS_SimpleBlur(blurMap1Sampler);
	} // pass GlowBlur2

	// And compose the final image with the Blurred Glow and the original image.
	pass ComposeFinalScene
	<
		string Script =
			"RenderColorTarget0=;"
			"Draw=Buffer;";        	
	>
	{
		//cullmode = none;
		//ZEnable = false;
		//ZWriteEnable = false;
		VertexShader = compile vs_1_1 VS_ScreenQuadSampleUp();
		PixelShader  = compile ps_1_1 PS_ComposeFinalImage(sceneMapSampler, blurMap2Sampler);
	} // pass ComposeFinalScene
} // technique ScreenGlow11
