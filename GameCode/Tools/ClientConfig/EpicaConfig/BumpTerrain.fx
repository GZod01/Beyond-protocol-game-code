uniform extern float4x4 gWorldInv : WorldInverse;
uniform extern float4x4 gWVP : WorldViewProjection;

float shininess : SpecularPower
<
    string UIWidget = "slider";
    float UIMin = 1.0;
    float UIMax = 128.0;
    float UIStep = 1.0;
    string UIName = "specular power";
> = 30.0;

float gfSpecMult
<
	string UIWidget = "slider";
	float UIMin = 0.1f;
	float UIMax = 2.0f;
	float UIStep = 0.05f;
	string UIName = "SpecMult";
> = 0.1f;

float4 lightDir : Direction
<
	string Object = "DirectionalLight";
    string Space = "World";
> = {0.58f, -0.58f, 0.58f, 0.0f};

float4 lightColor : Diffuse
<
    string UIName = "Diffuse Light Color";
    string Object = "DirectionalLight";
> = {1.0f, 1.0f, 1.0f, 1.0f};

float4 lightAmbient : Ambient
<
    string UIWidget = "Ambient Light Color";
    string Space = "material";
> = {0.19f, 0.19f, 0.19f, 1.0f};

float4 lightSpecular : Specular
<
	string UIWidget = "Light Specular";
	string Space = "material";
> = {1.0f, 1.0f, 1.0f, 1.0f};

float gfHalfSize = 24000.0f;
float gfFullSize = 48000.0f;
 
float gFogStart = 1.0f;
float gFogRange = 200.0f;
float4 gFogColor = {0.5f, 0.5f, 0.5f, 1.0f};
bool gbRenderFog;

uniform extern float3   gEyePosW : CameraPosition;
uniform extern texture  gTex;
uniform extern texture  gNormalMap;
uniform extern texture  gIllumMap;
uniform extern texture  gFOWMap;

sampler TexS = sampler_state
{
	Texture = <gTex>;
	MinFilter = ANISOTROPIC;
	MaxAnisotropy = 8;
	MagFilter = LINEAR;
	MipFilter = LINEAR;
	AddressU  = WRAP;
    AddressV  = WRAP;
    AddressW  = CLAMP;
};

sampler NormalMapS = sampler_state
{
	Texture = <gNormalMap>;
	MinFilter = ANISOTROPIC;
	MaxAnisotropy = 8;
	MagFilter = ANISOTROPIC;//LINEAR;
	MipFilter = ANISOTROPIC;//NONE;
	AddressU  = WRAP;
    AddressV  = WRAP;
    AddressW  = CLAMP;
};

sampler IllumMapS = sampler_state
{
	Texture = <gIllumMap>;
	MinFilter = ANISOTROPIC;
	MaxAnisotropy = 8;
	MagFilter = LINEAR;
	MipFilter = LINEAR;
	AddressU  = WRAP;
    AddressV  = WRAP;
};

sampler FogOfWarMap = sampler_state
{
	Texture = <gFOWMap>;
	MinFilter = ANISOTROPIC;
	MaxAnisotropy = 8;
	MagFilter = LINEAR;
	MipFilter = LINEAR;
	AddressU  = WRAP;
    AddressV  = WRAP;
    AddressW  = CLAMP;
};
 
struct outNormalMap
{
	float4 posH      : POSITION0;
    float3 toEyeT    : TEXCOORD0;
    float3 lightDirT : TEXCOORD1;
    float3 tex0      : TEXCOORD2;  
    float fogLerpParam : TEXCOORD3;
    float2 tex1		 : TEXCOORD4;
};

outNormalMap NormalMapVS(float3 posL      : POSITION0, 
                     float3 normalL   : NORMAL0, 
                     float3 tex0      : TEXCOORD0,
                     float3 tangentL  : TANGENT0,
                     float3 binormalL : BINORMAL0)
{
    // Zero out our output.
	outNormalMap outVS = (outNormalMap)0;
	
	// Build TBN-basis.
	float3x3 TBN;
	TBN[0] = tangentL;
	TBN[1] = binormalL;
	TBN[2] = normalL;
	
	// Matrix transforms from object space to tangent space.
	float3x3 toTangentSpace = transpose(TBN);
	
	// Transform eye position to local space.
	float3 eyePosL = mul(float4(gEyePosW, 1.0f), gWorldInv);
	
	// Transform to-eye vector to tangent space.
	float3 toEyeL = eyePosL - posL;
	outVS.toEyeT = mul(toEyeL, toTangentSpace);
	
	// Transform light direction to tangent space.
	float4 tmpLightDir = float4(lightDir.xyz, 0.0f);
	float3 lightDirL = mul(tmpLightDir, gWorldInv); //mul(lightDir, gWorldInv).xyz;
	outVS.lightDirT  = mul(lightDirL, toTangentSpace);
	
	// Transform to homogeneous clip space.
	//float4 posH = mul(float4(posL, 1.0f), gWVP); 
	outVS.posH = mul(float4(posL, 1.0f), gWVP);
		
	// Pass on texture coordinates to be interpolated in rasterization.
	//outVS.tex0 = tex0;
	outVS.tex0 = tex0;//float3(tex0.xy, 2.0f);
	//outVS.tex1 = tex1;
	//MSC - 06/15/08 - remarked out second texture coordinate and cleared the passed vertex parm in favor of calculating here
	//DX was not passing the parm correctly even though the values were identical... go M$
	outVS.tex1 = float3((posL.x + gfHalfSize) / gfFullSize, 1.0f - ((posL.z + gfHalfSize) / gfFullSize), 0.0f);
	
	//do our fog layer
	if (gbRenderFog)
	{
		float dist = distance(posL, eyePosL);
		outVS.fogLerpParam = saturate((dist - gFogStart) / gFogRange);
	}
		 	
	// Done--return the output.
    return outVS;
} 
 
float4 NormalMapFullPS(float4 posH : POSITION0,
				   float3 toEyeT    : TEXCOORD0,
                   float3 lightDirT : TEXCOORD1,
                   float3 tex0      : TEXCOORD2,
                   float fogLerpParam : TEXCOORD3,
                   float2 tex1		: TEXCOORD4) : COLOR
{
	// Interpolated normals can become unnormal--so normalize.
	toEyeT    = normalize(toEyeT);
	lightDirT = normalize(lightDirT);
	
	// Light vector is opposite the direction of the light.
	float3 lightVecT = -lightDirT;
	
	// Sample normal map.
	float3 normalT = tex3D(NormalMapS, tex0);//tex2D(NormalMapS, tex0);
		
	// Expand from [0, 1] compressed interval to true [-1, 1] interval.
    normalT = 2.0f*normalT - 1.0f;
            
    // Make it a unit vector.
	normalT = normalize(normalT);
	
	// Compute the reflection vector.
	float3 r = reflect(-lightVecT, normalT);

	// Determine how much (if any) specular light makes it into the eye.
	//float t  = pow(max(dot(r, toEyeT), 0.0f), shininess * 0.5f);
		
	// Determine the diffuse light intensity that strikes the vertex.
	float s = max(dot(lightVecT, normalT), 0.0f);
	
	//t = t * s * 0.5f; //* gfSpecMult;		//using this isnt a bad result either
		     
	float3 diffuse = s * lightColor.rgb;
	
	// Get the texture color.
	float4 texColor = tex3D(TexS, tex0) * 2.5f;	//1.25f
		
	float4 FOWColor = tex2D(FogOfWarMap, tex1);
	
	//float3 spec = t * lightSpecular.rgb;
		
	// Combine the color from lighting with the texture color.
	//float3 f3color = ( (((lightAmbient + diffuse)*texColor.rgb + (t*texColor)) ) * (FOWColor)) ; //+ (illumColor * texColor);
	//float3 f3color = ( (((lightAmbient + diffuse)*texColor.rgb ) ) + (FOWColor.rgb * texColor.rgb)) ;
	//float3 f3color = ( (((lightAmbient + diffuse)*texColor.rgb ) ) + (FOWColor.rgb * texColor.rgb)) ;
	//MSC - 06/12/08 - fix for fow texturing - still a little dark but gets the job done
	float3 f3color = ((lightAmbient + diffuse) * FOWColor.rgb) * texColor.rgb;
	
	//float3 f3color = (texColor.rgb * FOWColor.rgb) * (diffuse + lightAmbient);// * 0.5f);
						    
	if (tex0.z < 0.125f)
	{
		float t = 1.0f - (tex0.z / 0.125f);
		f3color = max(f3color , tex2D(IllumMapS, tex0 ) * t ) ;//* 0.5f;
	}
	
	//MSC 04/09/08 - Added fog into this
    float4 f4Result;
    if (gbRenderFog)
    {
    	f4Result = float4(lerp(f3color, gFogColor, fogLerpParam), texColor.a);
    }
    else
    {
    	f4Result = float4(f3color, texColor.a);
    }
    return f4Result;
}

float4 NormalMapOnlyPS(float4 posH : POSITION0,
				   float3 toEyeT    : TEXCOORD0,
                   float3 lightDirT : TEXCOORD1,
                   float3 tex0      : TEXCOORD2,
                   float fogLerpParam : TEXCOORD3,
                   float2 tex1		: TEXCOORD4) : COLOR
{
	// Interpolated normals can become unnormal--so normalize.
	toEyeT    = normalize(toEyeT);
	lightDirT = normalize(lightDirT);
	
	// Light vector is opposite the direction of the light.
	float3 lightVecT = -lightDirT;
	
	// Sample normal map.
	float3 normalT = tex3D(NormalMapS, tex0);
	
	// Expand from [0, 1] compressed interval to true [-1, 1] interval.
    normalT = 2.0f*normalT - 1.0f;
            
    // Make it a unit vector.
	normalT = normalize(normalT);
	
	// Compute the reflection vector.
	float3 r = reflect(-lightVecT, normalT);

	// Determine how much (if any) specular light makes it into the eye.
	float t  = pow(max(dot(r, toEyeT), 0.0f), shininess * 0.5f);
		
	// Determine the diffuse light intensity that strikes the vertex.
	float s = max(dot(lightVecT, normalT), 0.0f);
	
	t = t * s * gfSpecMult;		//using this isnt a bad result either
		     
	float3 diffuse = s * lightColor.rgb;
	
	// Get the texture color.
	float4 texColor = tex3D(TexS, tex0)*1.25f;
		
	float4 FOWColor = (tex2D(FogOfWarMap, tex1)) ;//* 0.55f) + 0.45f;
	
	// Combine the color from lighting with the texture color.
	//float3 f3color =( (((lightAmbient + diffuse)*texColor.rgb + (t*texColor)) ) * (FOWColor)) ; //+ (illumColor * texColor);
	//MSC - 06/12/08 - fix for fow texturing - still a little dark but gets the job done
	float3 f3color = ((lightAmbient + diffuse) * FOWColor.rgb) * texColor.rgb;
	//float3 f3color = (texColor.rgb * FOWColor.rgb) * (diffuse + lightAmbient);
    	
    //MSC 04/09/08 - Added fog into this
    float4 f4Result;
    if (gbRenderFog)
    {
    	f4Result = float4(lerp(f3color, gFogColor, fogLerpParam), texColor.a);
    }
    else
    {
    	f4Result = float4(f3color, texColor.a);
    }
    return f4Result;
}

 

technique NormalMapOnly
{
    pass P0
    {
        // Specify the vertex and pixel shader associated with this pass.
        vertexShader = compile vs_1_1 NormalMapVS();
        pixelShader  = compile ps_2_0 NormalMapOnlyPS();
    }
} 

technique NormalMapFull
{
    pass P0
    {
        // Specify the vertex and pixel shader associated with this pass.
        vertexShader = compile vs_1_1 NormalMapVS();
        pixelShader  = compile ps_2_0 NormalMapFullPS();
    }
}
 
 