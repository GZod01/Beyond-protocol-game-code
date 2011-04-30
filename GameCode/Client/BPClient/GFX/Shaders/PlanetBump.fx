//=============================================================================
// Does per-pixel lighting where the normals are obtained from a normal map.
//=============================================================================

uniform extern float4x4 gWorldInv : WorldInverse;
uniform extern float4x4 gWVP : WorldViewProjection;

float shininess : SpecularPower
<
    string UIWidget = "slider";
    float UIMin = 1.0;
    float UIMax = 128.0;
    float UIStep = 1.0;
    string UIName = "specular power";
> = 3.0;
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

uniform extern float3   gEyePosW : CameraPosition;
uniform extern texture  gTex;
uniform extern texture  gNormalMap;
uniform extern texture  gIllumMap; 

sampler TexS = sampler_state
{
	Texture = <gTex>;
	MinFilter = LINEAR;
	MagFilter = LINEAR;
	MipFilter = LINEAR;
};

sampler NormalMapS = sampler_state
{
	Texture = <gNormalMap>;
	MinFilter = LINEAR;
	MagFilter = LINEAR;
	MipFilter = LINEAR;
};

sampler IllumMapS = sampler_state
{
	Texture = <gIllumMap>;
	MinFilter = LINEAR;
	MagFilter = LINEAR;
	MipFilter = LINEAR;
};
 
struct OutputVS
{
    float4 posH      : POSITION0;
    float3 toEyeT    : TEXCOORD0;
    float3 lightDirT : TEXCOORD1;
    float2 tex0      : TEXCOORD2;
};

OutputVS NormalMapVS(float3 posL      : POSITION0, 
                     float3 normalL   : NORMAL0, 
                     float2 tex0      : TEXCOORD0,
                     float3 tangentL  : TANGENT0,
                     float3 binormalL : BINORMAL0)
{
    // Zero out our output.
	OutputVS outVS = (OutputVS)0;
	
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
	outVS.posH = mul(float4(posL, 1.0f), gWVP);
	
	// Pass on texture coordinates to be interpolated in rasterization.
	outVS.tex0 = tex0;	
		
	// Done--return the output.
    return outVS;
} 

float4 NormalMapPS(float3 toEyeT    : TEXCOORD0,
                   float3 lightDirT : TEXCOORD1,
                   float2 tex0      : TEXCOORD2) : COLOR
{
	// Interpolated normals can become unnormal--so normalize.
	toEyeT    = normalize(toEyeT);
	lightDirT = normalize(lightDirT);
	
	// Light vector is opposite the direction of the light.
	float3 lightVecT = -lightDirT;
				
	// Sample normal map.
	float3 normalT = tex2D(NormalMapS, tex0);
	
	// Expand from [0, 1] compressed interval to true [-1, 1] interval.
    normalT = 2.0f*normalT - 1.0f;
            
    // Make it a unit vector.
	normalT = normalize(normalT);
	
	// Compute the reflection vector.
	float3 r = reflect(-lightVecT, normalT);

	// Determine how much (if any) specular light makes it into the eye.
	//float t  = max(dot(r, toEyeT), 0.0f);//pow(max(dot(r, toEyeT), 0.0f), shininess * 0.5f);
	float t  =  max(dot(r, toEyeT), 0.0f) * shininess;
		
	// Determine the diffuse light intensity that strikes the vertex.
	float s = max(dot(lightVecT, normalT), 0.0f);
	
	// If the diffuse light intensity is low, kill the specular lighting term.
	// It doesn't look right to add specular light when the surface receives 
	// little diffuse light.
	t = t * s * 0.1f;		//using this isnt a bad result either
	     
	// Compute the ambient, diffuse and specular terms separatly. 
	//float3 spec = t*(materialSpecular*lightSpecular).rgb;
	float3 f3diffuse = (s*lightColor.rgb);
	
	// Get the texture color.
	float4 texColor = tex2D(TexS, tex0)*1.25f;
	float4 illumColor = tex2D(IllumMapS, tex0);
	
	// Combine the color from lighting with the texture color.
	float3 f3color = ((lightAmbient + f3diffuse)*texColor.rgb + t ) + (texColor * illumColor);
	
	// Output the color and the alpha.
    return float4(f3color, texColor.a);
} 

technique NormalMapTech
{
    pass P0
    { 
    	WRAP2=COORD0;
    	
        vertexShader = compile vs_1_1 NormalMapVS();
        pixelShader  = compile ps_2_0 NormalMapPS();
    }
}