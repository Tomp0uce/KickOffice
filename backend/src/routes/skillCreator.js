import { Router } from 'express';
import { chatCompletion } from '../services/llmClient.js';
import { logAndRespond } from '../utils/http.js';
import { SKILL_CREATOR_SYSTEM_PROMPT } from '../config/skillCreatorPrompt.js';
import { models } from '../config/models.js';

const skillCreatorRouter = Router();

const VALID_HOSTS = ['word', 'excel', 'powerpoint', 'outlook', 'all'];
const VALID_MODES = ['immediate', 'draft', 'agent'];
const ENDPOINT = 'POST /api/skill-creator';

/**
 * POST /api/skill-creator
 * Body: { description: string, host?: SkillHost }
 * Returns: SkillCreatorResponse (JSON)
 *
 * Uses the standard model tier for generation.
 * Retries once on JSON parse failure.
 */
skillCreatorRouter.post('/', async (req, res) => {
  const { description, host } = req.body;

  if (!description || typeof description !== 'string' || description.trim().length < 5) {
    return logAndRespond(
      res,
      400,
      { error: 'description is required (min 5 characters)' },
      ENDPOINT,
    );
  }
  if (description.trim().length > 2000) {
    return logAndRespond(
      res,
      400,
      { error: 'description must be at most 2000 characters' },
      ENDPOINT,
    );
  }
  if (host && !VALID_HOSTS.includes(host)) {
    return logAndRespond(
      res,
      400,
      { error: `host must be one of: ${VALID_HOSTS.join(', ')}` },
      ENDPOINT,
    );
  }

  const userMessage = [
    `L'utilisateur veut créer un skill pour : ${description.trim()}`,
    host && host !== 'all'
      ? `Host cible : ${host}`
      : 'Détermine le host approprié depuis la description.',
  ].join('\n');

  let lastError = null;

  // Try up to 2 times in case the LLM wraps the JSON in markdown or returns malformed JSON
  for (let attempt = 1; attempt <= 2; attempt++) {
    try {
      const modelConfig = models.standard;
      const body = {
        model: modelConfig.id,
        stream: false,
        max_tokens: 2048,
        temperature: 0.4,
        messages: [
          { role: 'system', content: SKILL_CREATOR_SYSTEM_PROMPT },
          {
            role: 'user',
            content:
              attempt === 1
                ? userMessage
                : `${userMessage}\n\nIMPORTANT: Réponds UNIQUEMENT avec l'objet JSON, sans aucun texte autour.`,
          },
        ],
      };

      const response = await chatCompletion({
        body,
        userCredentials: req.userCredentials,
        modelTier: 'standard',
      });

      if (!response.ok) {
        const errText = await response.text().catch(() => '');
        throw new Error(`LLM API error ${response.status}: ${errText.slice(0, 200)}`);
      }

      const data = await response.json();
      const rawText = data?.choices?.[0]?.message?.content || '';

      // Strip markdown code fences if present: ```json ... ```
      const stripped = rawText
        .replace(/^```(?:json)?\s*/i, '')
        .replace(/\s*```\s*$/, '')
        .trim();

      const parsed = JSON.parse(stripped);

      // Validate and sanitize required fields
      if (!parsed.name || !parsed.skillContent) {
        throw new Error('Missing required fields: name or skillContent');
      }
      if (!VALID_HOSTS.includes(parsed.host)) parsed.host = host || 'all';
      if (!VALID_MODES.includes(parsed.executionMode)) parsed.executionMode = 'immediate';
      parsed.icon = typeof parsed.icon === 'string' ? parsed.icon : 'Zap';
      parsed.description =
        typeof parsed.description === 'string' ? parsed.description : parsed.name;

      req.logger.info(`${ENDPOINT} skill created`, {
        skillName: parsed.name,
        host: parsed.host,
        executionMode: parsed.executionMode,
        traffic: 'user',
      });

      return res.json(parsed);
    } catch (err) {
      lastError = err;
      req.logger.warn(`${ENDPOINT} attempt ${attempt}/2 failed`, {
        error: err.message,
        traffic: 'system',
      });
    }
  }

  req.logger.error(`${ENDPOINT} failed after 2 attempts`, {
    error: lastError?.message,
    traffic: 'system',
  });
  return logAndRespond(
    res,
    500,
    { error: 'Failed to generate skill. Please try again.' },
    ENDPOINT,
  );
});

export { skillCreatorRouter };
