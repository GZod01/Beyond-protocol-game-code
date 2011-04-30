//=============================================================================
// Does per-pixel lighting where the normals are obtained from a normal map.
//=============================================================================

uniform extern float4x4 gWorldInv : WorldInverse;
uniform extern float4x4 gWVP : WorldViewProjection;
uniform extern float4x4 gWorld : World;
float4x4 worldInverseTranspose : WorldInverseTranspose;

//MSC - 10/25/08 - No longer used due to the new multi-tex setup
//float4 EntityColorAdjust : Diffuse
//<
//    string UIName = "Entity Color";
//> = {1.0f, 1.0f, 1.0f, 1.0f};

float4 RelColor
<
    string UIName = "Relationship Color";
> = {0.0f, 1.0f, 0.0f, 1.0f};

float shininess : SpecularPower
<
    string UIWidget = "slider";
    float UIMin = 1.0;
    float UIMax = 128.0;
    float UIStep = 1.0;
    string UIName = "specular power";
> = 30.0;
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

float gFogStart = 1.0f;
float gFogRange = 200.0f;
float4 gFogColor = {0.5f, 0.5f, 0.5f, 1.0f};
bool gbRenderFog;

uniform extern float3   gEyePosW : CameraPosition;
uniform extern texture  gTex;
uniform extern texture  gNormalMap;
uniform extern texture  gIllumMap;
uniform extern texture	gSpecMap;		//MSC - 10/25/08

sampler TexS = sampler_state
{
	Texture = <gTex>;
	MinFilter = ANISOTROPIC;
	MaxAnisotropy = 8;
	MagFilter = LINEAR;
	MipFilter = LINEAR;
	AddressU  = WRAP;
    AddressV  = WRAP;
};

sampler NormalMapS = sampler_state
{
	Texture = <gNormalMap>;
	MinFilter = ANISOTROPIC;
	MaxAnisotropy = 8;
	MagFilter = LINEAR;
	MipFilter = NONE;
	AddressU  = WRAP;
    AddressV  = WRAP;
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

//MSC - 10/25/08 - new specular component
sampler SpecMapS = sampler_state
{
	Texture = <gSpecMap>;
	MinFilter = ANISOTROPIC;
	MaxAnisotropy = 8;
	MagFilter = LINEAR;
	MipFilter = LINEAR;
	AddressU  = WRAP;
    AddressV  = WRAP;
};

struct outNormalMap
{
    float4 posH      : POSITION0;
    float3 toEyeT    : TEXCOORD0;
    float3 lightDirT : TEXCOORD1;
    float2 tex0      : TEXCOORD2;  
    float fogLerpParam : TEXCOORD3;
};

outNormalMap NormalMapVS(float3 posL      : POSITION0, 
                     float3 normalL   : NORMAL0, 
                     float2 tex0      : TEXCOORD0,
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
	float4 posH = mul(float4(posL, 1.0f), gWVP); 
	//outVS.posH = mul(float4(posL, 1.0f), gWVP);
	outVS.posH = posH;
	
	// Pass on texture coordinates to be interpolated in rasterization.
	outVS.tex0 = tex0;
	
	//do our fog layer
	if (gbRenderFog)
	{
		float dist = distance(posL, eyePosL);
		outVS.fogLerpParam = saturate((dist - gFogStart) / gFogRange);
	}
		 	
	// Done--return the output.
    return outVS;
} 

float4 NormalMapPS(float3 toEyeT    : TEXCOORD0,
                   float3 lightDirT : TEXCOORD1,
                   float2 tex0      : TEXCOORD2,
                   float fogLerpParam : TEXCOORD3) : COLOR
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
	float t  = pow(max(dot(r, toEyeT), 0.0f), shininess * 0.5f);
		
	// Determine the diffuse light intensity that strikes the vertex.
	float s = max(dot(lightVecT, normalT), 0.0f);
	
	// If the diffuse light intensity is low, kill the specular lighting term.
	// It doesn't look right to add specular light when the surface receives 
	// little diffuse light.
	//changing this to .5001 isn't a bad result
	//if(s <= 0.0001f)		
	//     t = 0.0f;
	t = t * s;		//using this isnt a bad result either
	     
	//s = 0.0f;
	
	// Compute the ambient, diffuse and specular terms separatly. 
	//MSC 04/07/08 - remarked out to remove the material data...
	float3 spec = t * lightSpecular.rgb;
	float3 diffuse = s * lightColor.rgb;

	// Get the texture color.
	float4 texColor = tex2D(TexS, tex0)*1.25f;
	
	// Combine the color from lighting with the texture color.
	float3 f3color = ((lightAmbient + diffuse)*(texColor).rgb + spec);// + (texColor * illumColor);
		
	// Output the color and the alpha.
	//MSC 04/07/08 - remarked out to remove material
    //return float4(f3color, materialDiffuse.a*texColor.a);
        
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

float4 NormalMapAndIllumPS(float3 toEyeT    : TEXCOORD0,
                   float3 lightDirT : TEXCOORD1,
                   float2 tex0      : TEXCOORD2,
                   float fogLerpParam : TEXCOORD3) : COLOR
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
	float t  = pow(max(dot(r, toEyeT), 0.0f), shininess * 0.5f);
		
	// Determine the diffuse light intensity that strikes the vertex.
	float s = max(dot(lightVecT, normalT), 0.0f);
	
	// If the diffuse light intensity is low, kill the specular lighting term.
	// It doesn't look right to add specular light when the surface receives 
	// little diffuse light.
	//changing this to .5001 isn't a bad result
	//if(s <= 0.0001f)		
	//     t = 0.0f;
	t = t * s;		//using this isnt a bad result either
	     
	//s = 0.0f;
	
	// Compute the ambient, diffuse and specular terms separatly. 
	//MSC 04/07/08 - remarked out to remove the material data...
	//float3 spec = t*(materialSpecular*lightSpecular).rgb;
	//float3 diffuse = s*(materialDiffuse*lightColor).rgb;
	//float3 ambient = materialAmbient*lightAmbient;
	float3 spec = t * lightSpecular.rgb;
	float3 diffuse = s * lightColor.rgb;
	//float3 ambient = materialAmbient*lightAmbient;
	
	// Get the texture color.
	float4 texColor = tex2D(TexS, tex0)*1.25f;
	float4 illumColor = tex2D(IllumMapS, tex0);
	
	// Combine the color from lighting with the texture color.
	float3 f3color = ((lightAmbient + diffuse)*(texColor).rgb + spec) + (texColor * illumColor)+(illumColor*RelColor);//(texColor * illumColor * RelColor);
		
	// Output the color and the alpha.
	//MSC 04/07/08 - remarked out to remove material
    //return float4(f3color, materialDiffuse.a*texColor.a);
    
    //MSC 04/09/08 - Added fog into this
    //return float4(f3color, texColor.a);
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

float4 NormalMapWithSpecularMapPS(float3 toEyeT    : TEXCOORD0,
                   float3 lightDirT : TEXCOORD1,
                   float2 tex0      : TEXCOORD2,
                   float fogLerpParam : TEXCOORD3) : COLOR
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
	float t  = pow(max(dot(r, toEyeT), 0.0f), shininess * 0.5f);
		
	// Determine the diffuse light intensity that strikes the vertex.
	float s = max(dot(lightVecT, normalT), 0.0f);
	
	// If the diffuse light intensity is low, kill the specular lighting term.
	// It doesn't look right to add specular light when the surface receives 
	// little diffuse light.
	//changing this to .5001 isn't a bad result
	//if(s <= 0.0001f)		
	//     t = 0.0f;
	t = t * s;		//using this isnt a bad result either
	     
	//s = 0.0f;
	
	// Compute the ambient, diffuse and specular terms separatly. 
	//MSC 04/07/08 - remarked out to remove the material data...
	float4 specMap = tex2D(SpecMapS, tex0) * 1.25f;
	float3 spec = t * lightSpecular.rgb * specMap * 4.0f;
	float3 diffuse = s * lightColor.rgb;

	// Get the texture color.
	float4 texColor = tex2D(TexS, tex0)*1.25f;
	
	// Combine the color from lighting with the texture color.
	float3 f3color = ((lightAmbient + diffuse)*(texColor * specMap).rgb + spec);// + (texColor * illumColor);
		
	// Output the color and the alpha.
	//MSC 04/07/08 - remarked out to remove material
    //return float4(f3color, materialDiffuse.a*texColor.a);
        
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

float4 NormalMapFullPS(float3 toEyeT    : TEXCOORD0,
                   float3 lightDirT : TEXCOORD1,
                   float2 tex0      : TEXCOORD2,
                   float fogLerpParam : TEXCOORD3) : COLOR
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
	float t  = pow(max(dot(r, toEyeT), 0.0f), shininess * 0.5f);
		
	// Determine the diffuse light intensity that strikes the vertex.
	float s = max(dot(lightVecT, normalT), 0.0f);
		
	// If the diffuse light intensity is low, kill the specular lighting term.
	// It doesn't look right to add specular light when the surface receives 
	// little diffuse light.
	//changing this to .5001 isn't a bad result
	//if(s <= 0.0001f)		
	//     t = 0.0f;
	t = t * s;		//using this isnt a bad result either
	     
	// Compute the ambient, diffuse and specular terms separatly. 
	// MSC - 07/22/09 - changed the shader to be a little better lit
	//float4 specMap = tex2D(SpecMapS, tex0) * 1.25f;
	//float3 spec = t * lightSpecular.rgb * specMap * 4.0f;
	float4 specMap = tex2D(SpecMapS, tex0) + 0.5f;
	float3 spec = t * lightSpecular.rgb * specMap * 2.0f;
	
	float3 diffuse = s * lightColor.rgb;		
	//float3 ambient = materialAmbient*lightAmbient;
	
	// Get the texture color.
	float4 texColor = tex2D(TexS, tex0)*1.25f;
	float4 illumColor = tex2D(IllumMapS, tex0);
	
	// Combine the color from lighting with the texture color.
	//MSC - 10/25/08
	//float3 f3color = ((lightAmbient + diffuse)*(texColor * EntityColorAdjust).rgb + spec) + (texColor * illumColor)+(illumColor*RelColor);// * RelColor);
	float3 f3color = ((lightAmbient + diffuse)*(texColor * specMap).rgb + spec) + (texColor * illumColor)+(illumColor*RelColor);// * RelColor);
		
	// Output the color and the alpha.
	//MSC 04/07/08 - remarked out to remove material
    //return float4(f3color, materialDiffuse.a*texColor.a);
    
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


struct outPerPixelVS {
    float4 posH				: POSITION;
    float4 texCoordDiffuse	: TEXCOORD0; 
    float3 normalW			: TEXCOORD1;
    float3 posW				: TEXCOORD2;
    float fogLerpParam		: TEXCOORD3;
};

outPerPixelVS PerPixelOnlyVS(float3 posL      : POSITION0, 
                     float3 normalL   : NORMAL0, 
                     float4 tex0      : TEXCOORD0) 
{
    outPerPixelVS OUT = (outPerPixelVS)0;
    float4 posH = mul( float4(posL, 1.0) , gWVP);
    OUT.posH = posH;
    OUT.texCoordDiffuse = tex0;
    
	OUT.normalW = normalize(mul(float4(normalL, 0.0f), worldInverseTranspose).xyz);
	OUT.posW = mul(float4(posL, 1.0f), gWorld).xyz;
		
	//do fog work
	if (gbRenderFog)
	{ 
		float3 eyePosL = mul(float4(gEyePosW, 1.0f), gWorldInv);
		float dist = distance(posH, eyePosL);
		OUT.fogLerpParam = saturate((dist - gFogStart) / gFogRange);
	}

    return OUT;
}

float4 PerPixelOnlyPS(outPerPixelVS oIN) : COLOR
{
	float3 toEye = normalize(gEyePosW - oIN.posW);
	float3 r = reflect(lightDir, oIN.normalW);
	float t = pow(max(dot(r, toEye), 0.0f), shininess * 0.5F);
	float s = max(dot(-lightDir, oIN.normalW), 0.0f);		//what's gLightVecW?
	
	//MSC - 04/07/08 - remarked out to get rid of the materialSpecular multiplier
	//float3 spec = t*(materialSpecular * lightColor).rgb;
	float3 spec = t * lightColor.rgb;
	float3 diffColor = tex2D( TexS, oIN.texCoordDiffuse );
	float3 diffuseTexture = s * diffColor;  // * (materialDiffuse * lightColor).rgb;
	
	//MSC - 04/07/08 - remarked out to get rid of the materialDiffuse, in the future, we may change 1.0f to be variable for the whole file
	//return float4(lightAmbient + spec + diffuseTexture , materialDiffuse.a);
	
	//MSC - 04/09/08 - added fog table
	float4 f4Result;
	if (gbRenderFog)
	{
		f4Result = float4( lerp(((lightAmbient + spec) * diffColor) + diffuseTexture, gFogColor, oIN.fogLerpParam), 1.0f); 
	}
	else
	{
		f4Result = float4(((lightAmbient + spec) * diffColor) + diffuseTexture, 1.0f); // * diffuseTexture , 1.0f);
	}
	return f4Result;
}

float4 PerPixelAndIllumPS(outPerPixelVS oIN) : COLOR
{
	float3 toEye = normalize(gEyePosW - oIN.posW);
	float3 r = reflect(lightDir, oIN.normalW);
	float t = pow(max(dot(r, toEye), 0.0f), shininess * 0.5F);
	float s = max(dot(-lightDir, oIN.normalW), 0.0f);		//what's gLightVecW?
	
	//MSC - 04/07/08 - remarked out to get rid of the materialSpecular multiplier
	//float3 spec = t*(materialSpecular * lightColor).rgb;
	float3 spec = t * lightColor.rgb;
	float3 diffColor = tex2D( TexS, oIN.texCoordDiffuse );
	float3 diffuseTexture = s * diffColor;  // * (materialDiffuse * lightColor).rgb;
	float3 v3Illum = tex2D ( IllumMapS, oIN.texCoordDiffuse ).rgb;
	
	//MSC - 04/07/08 - remarked out to get rid of the materialDiffuse, in the future, we may change 1.0f to be variable for the whole file
	//return float4(lightAmbient + spec + diffuseTexture , materialDiffuse.a);

	//MSC - 04/09/08 - added fog table
	float4 f4Result;
	if (gbRenderFog)
	{
		f4Result = float4( lerp(((lightAmbient + spec) * diffColor) + diffuseTexture + ((diffColor * v3Illum)+(v3Illum*RelColor)), gFogColor, oIN.fogLerpParam) , 1.0f);
	}
	else
	{
		f4Result = float4( ((lightAmbient + spec) * diffColor) + diffuseTexture + ((diffColor * v3Illum)+(v3Illum*RelColor)) , 1.0f);
	}	
	return f4Result;
}

//MSC - 10/25/08
float4 PerPixelWithSpecularMapPS(outPerPixelVS oIN) : COLOR
{
	float3 toEye = normalize(gEyePosW - oIN.posW);
	float3 r = reflect(lightDir, oIN.normalW);
	float t = pow(max(dot(r, toEye), 0.0f), shininess * 0.5F);
	float s = max(dot(-lightDir, oIN.normalW), 0.0f);		//what's gLightVecW?
	
	float4 specMap = tex2D(SpecMapS, oIN.texCoordDiffuse) * 1.25f;
	float3 spec = t * lightSpecular.rgb * specMap * 4.0f;
	float3 diffColor = tex2D( TexS, oIN.texCoordDiffuse ) * specMap;
	float3 diffuseTexture = s * diffColor;  // * (materialDiffuse * lightColor).rgb;
		
	float4 f4Result;
	if (gbRenderFog)
	{
		f4Result = float4( lerp(((lightAmbient + spec) * diffColor) + diffuseTexture, gFogColor, oIN.fogLerpParam), 1.0f); 
	}
	else
	{
		f4Result = float4(((lightAmbient + spec) * diffColor) + diffuseTexture, 1.0f); // * diffuseTexture , 1.0f);
	}
	return f4Result;
}

float4 PerPixelFullPS(outPerPixelVS oIN) : COLOR
{
	float3 toEye = normalize(gEyePosW - oIN.posW);
	float3 r = reflect(lightDir, oIN.normalW);
	float t = pow(max(dot(r, toEye), 0.0f), shininess * 0.5F);
	float s = max(dot(-lightDir, oIN.normalW), 0.0f);		//what's gLightVecW?
	
	//MSC - 04/07/08 - remarked out to get rid of the materialSpecular multiplier
	//float3 spec = t*(materialSpecular * lightColor).rgb;
	//MSC - 10/25/08
	//float3 spec = t * lightColor.rgb;
	float4 specMap = tex2D(SpecMapS, oIN.texCoordDiffuse) * 1.25f;
	float3 spec = t * lightSpecular.rgb * specMap * 4.0f;
	
	
	float3 diffColor = tex2D( TexS, oIN.texCoordDiffuse ) * specMap;//* EntityColorAdjust;	//MSC - 10/25/08
	float3 diffuseTexture = s * diffColor ;  // * (materialDiffuse * lightColor).rgb;
	float3 v3Illum = tex2D ( IllumMapS, oIN.texCoordDiffuse ).rgb;
		
	//MSC - 04/07/08 - remarked out to get rid of the materialDiffuse, in the future, we may change 1.0f to be variable for the whole file
	//return float4(lightAmbient + spec + diffuseTexture , materialDiffuse.a);
	
	//MSC - 04/09/08 - added fog table
	float4 f4Result;
	if (gbRenderFog)
	{
		f4Result = float4( lerp( ((lightAmbient + spec) * diffColor) + diffuseTexture + ((diffColor * v3Illum)+(v3Illum*RelColor)), gFogColor, oIN.fogLerpParam), 1.0f);
	}
	else
	{
		f4Result = float4( ((lightAmbient + spec) * diffColor) + diffuseTexture + ((diffColor * v3Illum)+(v3Illum*RelColor)), 1.0f);
	}
	return f4Result;
}

technique NormalMapOnly
{
    pass P0
    {
        // Specify the vertex and pixel shader associated with this pass.
        vertexShader = compile vs_1_1 NormalMapVS();
        pixelShader  = compile ps_2_0 NormalMapPS();
    }
}

technique NormalMapWithIllum
{
    pass P0
    {
        // Specify the vertex and pixel shader associated with this pass.
        vertexShader = compile vs_1_1 NormalMapVS();
        pixelShader  = compile ps_2_0 NormalMapAndIllumPS();
    }
}

//MSC - 10/25/08
technique NormalMapWithSpecularMap
{
    pass P0
    {
        // Specify the vertex and pixel shader associated with this pass.
        vertexShader = compile vs_1_1 NormalMapVS();
        pixelShader  = compile ps_2_0 NormalMapWithSpecularMapPS();
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

technique PerPixelOnly
{
    pass P0
    {
        // Specify the vertex and pixel shader associated with this pass.
        vertexShader = compile vs_1_1 PerPixelOnlyVS();
        pixelShader  = compile ps_2_0 PerPixelOnlyPS();
    }
}

technique PerPixelWithIllum
{
    pass P0
    {
        // Specify the vertex and pixel shader associated with this pass.
        vertexShader = compile vs_1_1 PerPixelOnlyVS();
        pixelShader  = compile ps_2_0 PerPixelAndIllumPS();
    }
}
//MSC - 10/25/08
technique PerPixelWithSpecularMap
{
    pass P0
    {
        // Specify the vertex and pixel shader associated with this pass.
        vertexShader = compile vs_1_1 PerPixelOnlyVS();
        pixelShader  = compile ps_2_0 PerPixelWithSpecularMapPS();
    }
}
technique PerPixelFull
{
    pass P0
    {
        // Specify the vertex and pixel shader associated with this pass.
        vertexShader = compile vs_1_1 PerPixelOnlyVS();
        pixelShader  = compile ps_2_0 PerPixelFullPS();
    }
}